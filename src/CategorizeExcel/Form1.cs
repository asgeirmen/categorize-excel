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
using System.IdentityModel.Tokens.Jwt;


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
            if (string.IsNullOrEmpty(textBoxSts.Text))
            {
                return null;
            }
            try
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
            }
            catch (Exception exception)
            {
                MessageBox.Show("Failed to get token: " + exception.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return null;
        }

        private string GetTapiXToken()
        {
            if (string.IsNullOrEmpty(textBoxSts.Text))
            {
                return null;
            }
            try
            {
                HttpClient cl = new HttpClient();
                var tokenResponse = cl.PostAsync($"{textBoxSts.Text}/auth/realms/tapix-prod/protocol/openid-connect/token",
                    new FormUrlEncodedContent(new Dictionary<string, string> {
                        { "grant_type", "client_credentials" },
                        { "scope", "user" },
                        { "client_id", textBoxClientId.Text },
                        { "client_secret", textBoxClientSecret.Text } })).Result;
                if (tokenResponse.IsSuccessStatusCode)
                {
                    var responseString = tokenResponse.Content.ReadAsStringAsync().Result;
                    var jsonObj = JsonObject.Parse(responseString);
                    var token = jsonObj["access_token"].ToString();
                    return token;
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show("Failed to get token: " + exception.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return null;
        }

        public static long GetTokenExpirationTime(string token)
        {
            var handler = new JwtSecurityTokenHandler();
            var jwtSecurityToken = handler.ReadJwtToken(token);
            var tokenExp = jwtSecurityToken.Claims.First(claim => claim.Type.Equals("exp")).Value;
            var ticks = long.Parse(tokenExp);
            return ticks;
        }

        public static bool CheckTokenIsValid(string token)
        {
            var tokenTicks = GetTokenExpirationTime(token);
            var tokenDate = DateTimeOffset.FromUnixTimeSeconds(tokenTicks).UtcDateTime;

            var now = DateTime.Now.ToUniversalTime();

            var valid = tokenDate >= now.AddSeconds(-30);

            return valid;
        }

        private bool IsTokenValid(string token)
        {
            JwtSecurityToken jwtSecurityToken;
            try
            {
                jwtSecurityToken = new JwtSecurityToken(token);
            }
            catch (Exception)
            {
                return false;
            }

            return jwtSecurityToken.ValidTo > DateTime.UtcNow.AddSeconds(-30);
        }

        private void InsertResultsColumnIfNotExists(string columnName)
        {
            var colInd = GetColumnIndex(columnName);
            if (colInd == -1)
            {
                dataGridViewExcel.Columns.Insert(0,
                    new DataGridViewTextBoxColumn() { Name = columnName, HeaderText = columnName });
            }
            else
            {
                foreach (var row in dataGridViewExcel.Rows)
                {
                    ((DataGridViewRow)row).Cells[colInd].Value = "";
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

        private string CategorizeRows(int startInd, int endInd, string token, string apiBase, string apiType)
        {
            JsonObject requestObj = new JsonObject();
            JsonArray transArray = new JsonArray();
            try
            {
                requestObj["context"] = JsonObject.Parse("{" + textBoxContext.Text + "}");
            }
            catch (Exception e)
            {
                return "Failed to parse context" + e.Message;
            }

            try
            {
                requestObj["options"] = JsonObject.Parse("{" + textBoxOptions.Text + "}");
            }
            catch (Exception e)
            {
                return "Failed to parse options" + e.Message;
            }

            try
            {
                requestObj["transactions"] = transArray;

                for (int rowInd = startInd; rowInd < endInd; rowInd++)
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
                            var standardFieldType = StandardFieldType(field);

                            if (standardFieldType != null)
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
                        row.Cells[GetColumnIndex("Request")].Value = JsonSerializer.Serialize(trans, new JsonSerializerOptions()
                        {
                            WriteIndented = false
                        });
                    }
                }

                if (transArray.Count == 0)
                {
                    return null;
                }
                var requestData = JsonSerializer.Serialize(requestObj, new JsonSerializerOptions());

                HttpClient cl = new HttpClient();
                if (token != null)
                {
                    cl.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);
                }

                string enrichEndpoint = apiType == "Core" ? $"{textBoxEnrichment.Text}/integration/v2/transactions/enrich" :
                    $"{apiBase}/integration/enrichment/v2/transactions/enrich";

                var apiResponse = cl.PostAsync(enrichEndpoint,
                    new StringContent(requestData, Encoding.UTF8, "application/json")).Result;
                if (apiResponse.IsSuccessStatusCode)
                {
                    var responseString = apiResponse.Content.ReadAsStringAsync().Result;
                    var transResults = JsonSerializer.Deserialize<JsonArray>(responseString);
                    for (int ind = 0; ind < transResults.Count; ind++)
                    {
                        var trans = (JsonObject)transResults[ind];
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

                        if (trans.ContainsKey("categoryId"))
                        {
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
                        }

                        if (trans.ContainsKey("normalizedText"))
                        {
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
                        }

                        if (trans.ContainsKey("displayText"))
                        {
                            row.Cells[GetColumnIndex("ResDisplayText")].Value = trans["displayText"].AsValue().ToString();
                        }

                        row.Cells[GetColumnIndex("Response")].Value = JsonSerializer.Serialize(trans, new JsonSerializerOptions()
                        {
                            WriteIndented = false
                        });
                    }
                }
                else
                {
                    return "Error reponse code " + apiResponse.Content.ReadAsStringAsync().Result;
                }


            }
            catch (Exception exception)
            {
                return "Error" + exception.Message;

            }

            return null;
        }

        private string? EnrichTapix(int startInd, int endInd, ref string? token, string apiBase)
        {
            JsonObject requestObj = new JsonObject();
            JsonArray transArray = new JsonArray();

            try
            {
                requestObj["requests"] = transArray;

                for (int rowInd = startInd; rowInd < endInd; rowInd++)
                {
                    var row = dataGridViewExcel.Rows[rowInd];

                    JsonObject trans = new JsonObject();
                    bool hasAnyColumn = false;
                    for (int ind = 0; ind < dataGridViewExcel.ColumnCount; ind++)
                    {
                        var colName = dataGridViewExcel.Columns[ind].Name;
                        var field = ToApiPropertyName(colName);
                        var colValue = ((DataGridViewRow)row).Cells[ind].Value;

                        if (colValue != null && colValue.ToString() != "")
                        {
                            trans[field] = JsonValue.Create(colValue);
                            hasAnyColumn = true;
                        }
                    }

                    if (hasAnyColumn)
                    {
                        transArray.Add(trans);
                        row.Cells[GetColumnIndex("Request")].Value = JsonSerializer.Serialize(trans, new JsonSerializerOptions()
                        {
                            WriteIndented = false
                        });
                    }

                }


                if (transArray.Count == 0)
                {
                    return null;
                }
                var requestData = JsonSerializer.Serialize(requestObj, new JsonSerializerOptions());
                HttpClient cl = new HttpClient();
                if (token != null)
                {
                    if (!IsTokenValid(token))
                    {
                        token = GetTapiXToken();
                    }

                    cl.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);
                }

                var apiResponse = cl.PostAsync(apiBase + "/v6/shops/findByCardTransactionBatch",
                    new StringContent(requestData, Encoding.UTF8, "application/json")).Result;
                if (apiResponse.IsSuccessStatusCode)
                {
                    var responseString = apiResponse.Content.ReadAsStringAsync().Result;
                    var transResults = JsonSerializer.Deserialize<JsonArray>(responseString);
                    for (int ind = 0; ind < transResults.Count; ind++)
                    {
                        var trans = (JsonObject)transResults[ind];
                        var row = (DataGridViewRow)dataGridViewExcel.Rows[startInd + ind];
                        row.Cells[GetColumnIndex("ResHandle")].Value = trans["handle"];
                        if (trans["shop"] != null)
                        {
                            var shopId = (trans["shop"]["uid"]).AsValue().ToString();
                            var shopResponse = cl.GetAsync(apiBase + "/v6/shops/" + shopId).Result;
                            if (shopResponse.IsSuccessStatusCode)
                            {
                                var shopResponseString = shopResponse.Content.ReadAsStringAsync().Result;
                                var shopResults = JsonSerializer.Deserialize<JsonObject>(shopResponseString);
                                trans["shop"] = shopResults;

                                if (shopResults["category"] != null && shopResults["category"]["name"] != null)
                                {
                                    row.Cells[GetColumnIndex("ResCategory")].Value = shopResults["category"]["name"].AsValue().ToString();
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

                                if (shopResults["location"] != null && shopResults["location"]["coordinates"] != null)
                                {
                                    var coordinates = shopResults["location"]["coordinates"];
                                    row.Cells[GetColumnIndex("ResLocation")].Value = coordinates["lat"].AsValue().ToString() + " " + coordinates["long"].AsValue().ToString();
                                }
                                if (shopResults["googlePlaceId"] != null)
                                {
                                    row.Cells[GetColumnIndex("ResPlaceId")].Value = shopResults["googlePlaceId"].AsValue().ToString();
                                }

                                if (shopResults["tags"] != null && shopResults["tags"].AsArray().FirstOrDefault(n => n.AsValue().ToString() == "Subscription") != null)
                                {
                                    row.Cells[GetColumnIndex("ResSubscription")].Value = true;
                                }

                                if (trans["shop"]["merchantUid"] != null)
                                {
                                    var merchantIdId = (trans["shop"]["merchantUid"]).AsValue().ToString();
                                    var merchantResponse = cl.GetAsync(apiBase + "/v6/merchants/" + merchantIdId).Result;
                                    if (shopResponse.IsSuccessStatusCode)
                                    {
                                        var merchantResponseString = merchantResponse.Content.ReadAsStringAsync().Result;
                                        var merchantResults = JsonSerializer.Deserialize<JsonObject>(merchantResponseString);

                                        if (merchantResults["logo"] != null)
                                        {
                                            row.Cells[GetColumnIndex("ResLogo")].Value = merchantResults["logo"].AsValue().ToString();
                                        }

                                        trans["merchant"] = merchantResults;
                                        row.Cells[GetColumnIndex("ResDisplayText")].Value = merchantResults["name"].AsValue().ToString();

                                    }
                                }
                            }

                        }

                        if (trans.ContainsKey("normalizedText"))
                        {
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
                        }

                        if (trans.ContainsKey("displayText"))
                        {
                            row.Cells[GetColumnIndex("ResDisplayText")].Value = trans["displayText"].AsValue().ToString();
                        }

                        row.Cells[GetColumnIndex("Response")].Value = JsonSerializer.Serialize(trans, new JsonSerializerOptions()
                        {
                            WriteIndented = false
                        });
                    }
                }
                else
                {
                    return "Error reponse code " + apiResponse.StatusCode.ToString() + ": " + apiResponse.Content.ReadAsStringAsync().Result;
                }

            }
            catch (Exception exception)
            {
                return "Error: " + exception.Message;

            }

            return null;
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

            InsertResultsColumnIfNotExists("ResHandle");
            InsertResultsColumnIfNotExists("ResLocation");
            InsertResultsColumnIfNotExists("ResLogo");
            InsertResultsColumnIfNotExists("ResPlaceId");
            InsertResultsColumnIfNotExists("ResSubscription");
            InsertResultsColumnIfNotExists("ResCategoryId");
            InsertResultsColumnIfNotExists("ResCategory");
            InsertResultsColumnIfNotExists("ResNormalizedText");
            InsertResultsColumnIfNotExists("ResDisplayText");
            InsertResultsColumnIfNotExists("Response");
            InsertResultsColumnIfNotExists("Request");

            FindSpecialColumnIds();

            var apiBase = textBoxApi.Text;
            var apiType = comboBoxApiType.SelectedItem.ToString();
            string? token;
            switch (apiType)
            {
                case "Enrichment":
                    token = GetToken();
                    break;
                case "TapiX":
                    token = GetTapiXToken();
                    break;
                default:
                    token = null;
                    break;
            }

            Thread backgroundThread = new Thread(
                new ThreadStart(() =>
                    {
                        int batchSize = 50;
                        for (int ind = 0; ind < dataGridViewExcel.RowCount;)
                        {
                            int endInd = Math.Min(dataGridViewExcel.RowCount, ind + batchSize);

                            string errorMsg = null;
                            switch (apiType)
                            {
                                case "Core":
                                case "Enrichment":
                                    errorMsg = CategorizeRows(ind, endInd, token, apiBase, apiType);
                                    break;
                                case "TapiX":
                                    errorMsg = EnrichTapix(ind, endInd, ref token, apiBase);
                                    break;
                            }

                            if (!string.IsNullOrEmpty(errorMsg))
                            {
                                
                                buttonCategorizeExcel.BeginInvoke(
                                    new Action(() =>
                                    {
                                        MessageBox.Show(errorMsg, "Enrichment failed: " + errorMsg, MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    }
                                    ));
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
                                    _categorizationInProgress = false;
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
                            if (StandardFieldType(field) == null)
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

        private Type? StandardFieldType(string fieldName)
        {
            switch (fieldName)
            {
                case "identifier":
                case "text":
                case "currency":
                case "counterpartyAccountId":
                case "counterpartyName":
                case "terminalId":
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
                    return typeof(string);
                case "transactionDate":
                case "bookingDate":
                case "valueDate":
                case "timestamp":
                    return typeof(DateTime);
                case "mcc":
                case "amount":
                case "amountInCurrency":
                case "bookedAmount":
                case "accountBalance":
                    return typeof(decimal);
                case "isMerchant":
                case "isOwnAccountTransfer":
                case "isPending":
                    return typeof(bool);
            }

            return null;
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
                if (columnName[charInd - 1] == '_' || columnName[charInd - 1] == ' ' || Char.IsLower(columnName[charInd - 1]) && Char.IsUpper(columnName[charInd]))
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

        private void comboBoxSheet_Format(object sender, ListControlConvertEventArgs e)
        {
            e.Value = ((DataRowView)e.Value).Row["TABLE_NAME"].ToString();
        }

    }
}