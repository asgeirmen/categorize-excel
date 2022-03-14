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

        public Form1()
        {
            InitializeComponent();
        }

        private string GetToken()
        {
            HttpClient cl = new HttpClient();
            string encoded = System.Convert.ToBase64String(Encoding.GetEncoding("ISO-8859-1")
                .GetBytes(textBoxClientId.Text + ":" + textBoxClientSecret.Text));
            cl.DefaultRequestHeaders.Add("Authorization", "Basic " + encoded);
            var tokenResponse = cl.PostAsync($"{textBoxSts.Text}/connect/token",
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
            try
            {

                JsonObject requestObj = new JsonObject();
                JsonArray transArray = new JsonArray();
                requestObj["transactions"] = transArray;
                requestObj["categoryContextId"] = JsonValue.Create(1);
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
                    bool hasAnyColumn = false;
                    for (int ind = 0; ind < dataGridViewExcel.ColumnCount; ind++)
                    {
                        var colName = dataGridViewExcel.Columns[ind].Name.ToLower();
                        var field = ToApiPropertyName(colName);
                        var colValue = ((DataGridViewRow)row).Cells[ind].Value;

                        if (colValue != null)
                        {
                            switch (field)
                            {
                                case "identifier":
                                case "text":
                                case "currency":
                                case "externalMerchantIdentifier":
                                case "countryCode":
                                case "date":
                                case "dueDate":
                                case "mcc":
                                case "amount":
                                case "amountInCurrency":
                                case "bookedAmount":
                                case "isMerchant":
                                case "isOwnAccountTransfer":
                                case "isUncleared":
                                    if (colValue.ToString() != "")
                                    {
                                        trans[field] = JsonValue.Create(colValue);
                                        hasAnyColumn = true;
                                    }
                                    break;
                            }
                        }
                    }

                    if (hasAnyColumn)
                    {
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

                var apiResponse = cl.PostAsync($"{textBoxApi.Text}/integration/core/v1/transactions/categorize/{textBoxAccountTypeProfile.Text}",
                    new StringContent(requestData, Encoding.UTF8, "application/json")).Result;
                if (apiResponse.IsSuccessStatusCode)
                {
                    var responseString = apiResponse.Content.ReadAsStringAsync().Result;
                    var reponseObject = JsonSerializer.Deserialize<JsonObject>(responseString);
                    var transResults = reponseObject["data"] as JsonArray;
                    for (int ind = 0; ind < transResults.Count; ind++)
                    {
                        var trans = transResults[ind];
                        var row = dataGridViewExcel.Rows[startInd + ind];
                        ((DataGridViewRow)row).Cells[GetColumnIndex("ResCategoryId")].Value = trans["categoryId"].AsValue().ToString();
                        ((DataGridViewRow)row).Cells[GetColumnIndex("ResSubText")].Value = trans["subText"].AsValue().ToString();
                        ((DataGridViewRow)row).Cells[GetColumnIndex("Response")].Value = JsonSerializer.Serialize(trans, new JsonSerializerOptions()
                        {
                            WriteIndented = true
                        });
                        ((DataGridViewRow)row).Cells[GetColumnIndex("Request")].Value = JsonSerializer.Serialize(transArray[ind], new JsonSerializerOptions()
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
            buttonCategorizeExcel.Enabled = false;
            progressBarCategorize.Value = 0;
            InsertResultsColumnIfNotExists("ResCategoryId");
            InsertResultsColumnIfNotExists("ResSubText");
            InsertResultsColumnIfNotExists("Response");
            InsertResultsColumnIfNotExists("Request");

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
                        }

                        buttonCategorizeExcel.BeginInvoke(
                            new Action(() => { buttonCategorizeExcel.Enabled = true; }
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
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failed to read Excel file: " + ex.Message, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error  

                }
            }

        }

        private string ToApiPropertyName(string claimName)
        {
            var strBuilder = new StringBuilder(claimName.Length);
            strBuilder.Append(claimName[0]);

            for (int charInd = 1; charInd < claimName.Length; charInd++)
            {
                char curChar = claimName[charInd];
                if (claimName[charInd - 1] == '_')
                {
                    strBuilder.Append(char.ToUpper(curChar));
                }
                else if (curChar != '_')
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
    }
}