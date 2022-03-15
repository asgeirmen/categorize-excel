using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Data;
using System.Net.Http.Json;
using System.Text.Json;
using System.Text.Json.Nodes;


namespace CategorizeExcel
{
    public partial class Form1 : Form
    {
        private string _excelFilePath = null;
        private int _categoryIdColumnIndex = -1;
        private int _categoryNameColumnIndex = -1;
        private int _normalizedTextColumnIndex = -1;
        private bool _categorizationInProgress = false;
        private bool _categorizationHasBeenCancelled = false;
        public Form1()
        {
            InitializeComponent();
            this.Icon = Properties.Resources.favicon;
        }

        private string GetToken()
        {
            HttpClient cl = new HttpClient();
            string encoded = System.Convert.ToBase64String(Encoding.GetEncoding("ISO-8859-1")
                .GetBytes(textBoxClientId.Text + ":" + textBoxClientSecret.Text));
            cl.DefaultRequestHeaders.Add("Authorization", "Basic " + encoded);
            var tokenResponse = cl.PostAsync($"{textBoxSts.Text}/identity/connect/token",
                new FormUrlEncodedContent(new Dictionary<string, string> { { "grant_type", "client_credentials" } })).Result;
            if (tokenResponse.IsSuccessStatusCode)
            {
                var responseString = tokenResponse.Content.ReadAsStringAsync().Result;
                var jsonObj = JsonObject.Parse(responseString);
                var token = jsonObj["access_token"].ToString();
                return token;
            }

            throw new Exception("Failed to get token");
        }

        private void InsertResultsColumnIfNotExists(string columnName)
        {
            var colInd = GetColumnIndex(columnName);
            if (colInd == -1)
            {
                dataGridViewExcel.Columns.Insert(0,
                    new DataGridViewTextBoxColumn() {Name = columnName, HeaderText = columnName});
            }
            else
            {
                foreach (var row in dataGridViewExcel.Rows)
                {
                    ((DataGridViewRow) row).Cells[colInd].Value = "";
                }
            }
        }

        private int GetColumnIndex(string columnName)
        {
            for (int ind = 0; ind < dataGridViewExcel.ColumnCount; ind++)
            {
                var colName = dataGridViewExcel.Columns[ind].Name;
                if (columnName == colName)
                {
                    return ind;
                }
            }

            return -1;
        }

        private bool CategorizeRows(int startInd, int endInd)
        {
            JsonObject requestObj = new JsonObject();
            JsonArray transArray = new JsonArray();
            try
            {
                requestObj["context"] = JsonObject.Parse("{" + textBoxContext.Text + "}");
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Failed to parse context", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            try
            {
                requestObj["options"] = JsonObject.Parse("{" + textBoxOptions.Text + "}");
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Failed to parse options", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            try
            {

                requestObj["transactions"] = transArray;
                JsonObject optionsObj = new JsonObject
                {
                    ["includeDetectedCategories"] = JsonValue.Create(true),
                    ["includeNormalizationInfo"] = JsonValue.Create(true),
                    ["includeMerchantInfo"] = JsonValue.Create(true),
                    ["includeCategorizationBreakdown"] = JsonValue.Create(true)
                };
                requestObj["options"] = optionsObj;

                for(int rowInd = startInd; rowInd < endInd; rowInd++)
                {
                    var row = dataGridViewExcel.Rows[rowInd];

                    JsonObject trans = new JsonObject();
                    JsonObject customFields = new JsonObject();
                    bool hasAnyColumn = false;
                    for (int ind = 0; ind < dataGridViewExcel.ColumnCount; ind++)
                    {
                        var colName = dataGridViewExcel.Columns[ind].Name;
                        var field = ToApiPropertyName(colName);
                        var colValue = ((DataGridViewRow)row).Cells[ind].Value;

                        if (colValue != null && colValue.ToString() != "")
                        {
                            if (IsStandardField(field))
                            {
                                trans[field] = JsonValue.Create(colValue);
                                hasAnyColumn = true;
                            }
                            else if (IsCustomField(field))
                            {
                                customFields[field] = JsonValue.Create(colValue);
                                hasAnyColumn = true;
                            }
                        }
                    }

                    if (hasAnyColumn)
                    {
                        if (customFields.Count > 0)
                        {
                            trans["customFields"] = customFields;
                        }
                        transArray.Add(trans);
                    }
                }

                if (transArray.Count == 0)
                {
                    return true;
                }
                var requestData = JsonSerializer.Serialize(requestObj);

                var token = GetToken();
                HttpClient cl = new HttpClient();
                cl.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);

                var apiResponse = cl.PostAsync($"{textBoxApi.Text}/integration/enrichment/v2/transactions/enrich",
                    new StringContent(requestData, Encoding.UTF8, "application/json")).Result;
                if (apiResponse.IsSuccessStatusCode)
                {
                    var responseString = apiResponse.Content.ReadAsStringAsync().Result;
                    var transResults = JsonSerializer.Deserialize<JsonArray>(responseString);
                    for (int ind = 0; ind < transResults.Count; ind++)
                    {
                        var trans = transResults[ind];
                        var row = (DataGridViewRow)dataGridViewExcel.Rows[startInd + ind];
                        if (trans["categoryDetails"] != null)
                        {
                            row.Cells[GetColumnIndex("ResCategory")].Value = (trans["categoryDetails"]["label"]).AsValue().ToString();
                            if (_categoryNameColumnIndex >= 0)
                            {
                                if (row.Cells[_categoryNameColumnIndex].Value?.ToString()?.Trim() ==
                                    row.Cells[GetColumnIndex("ResCategory")]?.Value.ToString()?.Trim())
                                {
                                    row.Cells[GetColumnIndex("ResCategory")].Style.BackColor = Color.Chartreuse;
                                }
                                else
                                {
                                    row.Cells[GetColumnIndex("ResCategory")].Style.BackColor = Color.LightCoral;
                                }
                            }
                        }

                        row.Cells[GetColumnIndex("ResCategoryId")].Value = trans["categoryId"].AsValue().ToString();
                        if (_categoryIdColumnIndex >= 0)
                        {
                            if (row.Cells[_categoryIdColumnIndex].Value?.ToString()?.Trim() ==
                                row.Cells[GetColumnIndex("ResCategoryId")]?.Value.ToString()?.Trim())
                            {
                                row.Cells[GetColumnIndex("ResCategoryId")].Style.BackColor = Color.Chartreuse;
                            }
                            else
                            {
                                row.Cells[GetColumnIndex("ResCategoryId")].Style.BackColor = Color.LightCoral;
                            }
                        }
                        row.Cells[GetColumnIndex("ResNormalizedText")].Value = trans["normalizedText"].AsValue().ToString();
                        if (_normalizedTextColumnIndex >= 0)
                        {
                            if (row.Cells[_normalizedTextColumnIndex].Value?.ToString()?.Trim() ==
                                row.Cells[GetColumnIndex("ResNormalizedText")]?.Value.ToString()?.Trim())
                            {
                                row.Cells[GetColumnIndex("ResNormalizedText")].Style.BackColor = Color.Chartreuse;
                            }
                            else
                            {
                                row.Cells[GetColumnIndex("ResNormalizedText")].Style.BackColor = Color.LightCoral;
                            }
                        }

                        row.Cells[GetColumnIndex("ResDisplayText")].Value = trans["displayText"].AsValue().ToString();
                        row.Cells[GetColumnIndex("Response")].Value = JsonSerializer.Serialize(trans, new JsonSerializerOptions()
                        {
                            WriteIndented = true
                        });
                        row.Cells[GetColumnIndex("Request")].Value = JsonSerializer.Serialize(transArray[ind], new JsonSerializerOptions()
                        {
                            WriteIndented = true
                        });
                    }
                }
                else
                {
                    MessageBox.Show(apiResponse.Content.ReadAsStringAsync().Result, "Error reponse code " + apiResponse.StatusCode.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }


            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;

            }

            return true;
        }

        private void buttonCategorizeExcel_Click(object sender, EventArgs e)
        {
            if (_categorizationInProgress)
            {
                _categorizationInProgress = false;
                buttonCategorizeExcel.Enabled = false;
                return;
            }

            buttonCategorizeExcel.Text = "Cancel";
            progressBarCategorize.Value = 0;
            _categorizationInProgress = true;

            InsertResultsColumnIfNotExists("ResCategoryId");
            InsertResultsColumnIfNotExists("ResCategory");
            InsertResultsColumnIfNotExists("ResNormalizedText");
            InsertResultsColumnIfNotExists("ResDisplayText");
            InsertResultsColumnIfNotExists("Response");
            InsertResultsColumnIfNotExists("Request");

            FindSpecialColumnIds();

            Thread backgroundThread = new Thread(
                new ThreadStart(() =>
                    {
                        int batchSize = 50;
                        for (int ind = 0; ind < dataGridViewExcel.RowCount; )
                        {
                            int endInd = Math.Min(dataGridViewExcel.RowCount, ind + batchSize);
                            if (!CategorizeRows(ind, endInd))
                            {
                                break;
                            }

                            ind = endInd;

                            progressBarCategorize.BeginInvoke(
                                new Action(() =>
                                    {
                                        progressBarCategorize.Value = (100 * ind / dataGridViewExcel.RowCount);
                                    }
                                ));

                            if (!_categorizationInProgress)
                            {
                                break;
                            }
                        }

                        buttonCategorizeExcel.BeginInvoke(
                            new Action(() =>
                                {
                                    buttonCategorizeExcel.Enabled = true;
                                    buttonCategorizeExcel.Text = "Categorize";
                                }
                            ));
                    }
                ));
            backgroundThread.Start();

        }

        public void ReadExcel(string fileName, bool loadSheets)
        {
            var fileExt = Path.GetExtension(fileName);
            string conn = string.Empty;
            if (fileExt.CompareTo(".xls") == 0)
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
            else
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=Yes';"; //for above excel 2007  
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                con.Open();
                try
                {
                    if (loadSheets)
                    {
                        var dbSchema = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                        comboBoxSheet.DataSource = dbSchema;
                    }
                    else
                    {
                        DataTable dtexcel = new DataTable();
                        OleDbDataAdapter oleAdpt = new OleDbDataAdapter($"select * from [{comboBoxSheet.Text}]", con); //here we read data from sheet1  
                        oleAdpt.Fill(dtexcel);
                        dataGridViewExcel.Visible = true;
                        dataGridViewExcel.DataSource = dtexcel;

                        checkedListBoxCustomFields.Items.Clear();
                        for (int ind = 0; ind < dataGridViewExcel.ColumnCount; ind++)
                        {
                            var colName = dataGridViewExcel.Columns[ind].Name;
                            var field = ToApiPropertyName(colName);
                            if (!IsStandardField(field))
                            {
                                checkedListBoxCustomFields.Items.Add(field);
                            }
                        }
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failed to read Excel file: " + ex.Message, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error  

                }
            }

        }

        private void FindSpecialColumnIds()
        {
            for (int ind = 0; ind < dataGridViewExcel.ColumnCount; ind++)
            {
                var colName = dataGridViewExcel.Columns[ind].Name;
                var propName = ToApiPropertyName(colName);
                if (propName.StartsWith("res") || propName.Contains("parent"))
                {
                    // Not special column
                }
                else if (propName.Contains("categoryId"))
                {
                    _categoryIdColumnIndex = ind;
                }
                else if (propName.Contains("category"))
                {
                    _categoryNameColumnIndex = ind;
                }
                else if (propName.Contains("normalized") || propName.Contains("cleaned"))
                {
                    _normalizedTextColumnIndex = ind;
                }

            }
        }

        private bool IsStandardField(string fieldName)
        {
            switch (fieldName)
            {
                case "identifier":
                case "text":
                case "currency":
                case "counterpartyAccountId":
                case "counterpartyName":
                case "TerminalId":
                case "externalMerchantId":
                case "merchantName":
                case "countryCode":
                case "city":
                case "street":
                case "postalCode":
                case "region":
                case "geoLocation":
                case "maskedPan":
                case "checkId":
                case "purposeCode":
                case "bankTransactionCode":
                case "creditorId":
                case "reference":
                case "transactionDate":
                case "bookingDate":
                case "valueDate":
                case "timestamp":
                case "mcc":
                case "amount":
                case "amountInCurrency":
                case "bookedAmount":
                case "accountBalance":
                case "isMerchant":
                case "isOwnAccountTransfer":
                case "isPending":
                    return true;
            }

            return false;
        }

        private bool IsCustomField(string fieldName)
        {
            foreach (var customField in checkedListBoxCustomFields.CheckedItems)
            {
                if (fieldName == customField.ToString())
                {
                    return true;
                }
            }

            return false;
        }

        private string ToApiPropertyName(string columnName)
        {
            columnName = columnName.Trim();
            var strBuilder = new StringBuilder(columnName.Length);
            strBuilder.Append(Char.ToLower(columnName[0]));

            for (int charInd = 1; charInd < columnName.Length; charInd++)
            {
                char curChar = columnName[charInd];
                if (columnName[charInd - 1] == '_' || columnName[charInd - 1] == ' ')
                {
                    strBuilder.Append(char.ToUpper(curChar));
                }
                else if (curChar != '_' && curChar != ' ')
                {
                    strBuilder.Append(char.ToLower(curChar));
                }
            }

            return strBuilder.ToString();
        }

        private void buttonFindFile_Click(object sender, EventArgs e)
        {
            string fileExt = string.Empty;
            OpenFileDialog file = new OpenFileDialog()
            {
                Filter = "Excel Files (*.xls;*.xlsx)|*.xls;*.xlsx"
            };
            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK) //if there is a file choosen by the user  
            {
                _excelFilePath = file.FileName; //get the path of the file  
                try
                {

                    ReadExcel(_excelFilePath, true); 
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }

            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void comboBoxSheet_SelectedValueChanged(object sender, EventArgs e)
        {
            ReadExcel(_excelFilePath, false);
        }

        private void comboBoxSheet_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBoxSheet_Format(object sender, ListControlConvertEventArgs e)
        {
            e.Value = ((DataRowView) e.Value).Row["TABLE_NAME"].ToString();
        }

        private void textBoxContext_TextChanged(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }
    }
}