namespace CategorizeExcel
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.buttonCategorizeExcel = new System.Windows.Forms.Button();
            this.buttonFindFile = new System.Windows.Forms.Button();
            this.dataGridViewExcel = new System.Windows.Forms.DataGridView();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label5 = new System.Windows.Forms.Label();
            this.checkedListBoxCustomFields = new System.Windows.Forms.CheckedListBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.textBoxOptions = new System.Windows.Forms.TextBox();
            this.textBoxContext = new System.Windows.Forms.TextBox();
            this.progressBarCategorize = new System.Windows.Forms.ProgressBar();
            this.comboBoxSheet = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.textBoxClientSecret = new System.Windows.Forms.TextBox();
            this.textBoxClientId = new System.Windows.Forms.TextBox();
            this.textBoxSts = new System.Windows.Forms.TextBox();
            this.textBoxApi = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewExcel)).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // buttonCategorizeExcel
            // 
            this.buttonCategorizeExcel.Location = new System.Drawing.Point(12, 72);
            this.buttonCategorizeExcel.Name = "buttonCategorizeExcel";
            this.buttonCategorizeExcel.Size = new System.Drawing.Size(150, 46);
            this.buttonCategorizeExcel.TabIndex = 1;
            this.buttonCategorizeExcel.Text = "Categorize";
            this.buttonCategorizeExcel.UseVisualStyleBackColor = true;
            this.buttonCategorizeExcel.Click += new System.EventHandler(this.buttonCategorizeExcel_Click);
            // 
            // buttonFindFile
            // 
            this.buttonFindFile.Location = new System.Drawing.Point(12, 16);
            this.buttonFindFile.Name = "buttonFindFile";
            this.buttonFindFile.Size = new System.Drawing.Size(150, 46);
            this.buttonFindFile.TabIndex = 2;
            this.buttonFindFile.Text = "Load excel...";
            this.buttonFindFile.UseVisualStyleBackColor = true;
            this.buttonFindFile.Click += new System.EventHandler(this.buttonFindFile_Click);
            // 
            // dataGridViewExcel
            // 
            this.dataGridViewExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridViewExcel.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewExcel.Location = new System.Drawing.Point(0, 199);
            this.dataGridViewExcel.Name = "dataGridViewExcel";
            this.dataGridViewExcel.RowHeadersWidth = 82;
            this.dataGridViewExcel.RowTemplate.Height = 41;
            this.dataGridViewExcel.Size = new System.Drawing.Size(2202, 902);
            this.dataGridViewExcel.TabIndex = 3;
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.checkedListBoxCustomFields);
            this.panel1.Controls.Add(this.label7);
            this.panel1.Controls.Add(this.label6);
            this.panel1.Controls.Add(this.textBoxOptions);
            this.panel1.Controls.Add(this.textBoxContext);
            this.panel1.Controls.Add(this.progressBarCategorize);
            this.panel1.Controls.Add(this.comboBoxSheet);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.textBoxClientSecret);
            this.panel1.Controls.Add(this.textBoxClientId);
            this.panel1.Controls.Add(this.textBoxSts);
            this.panel1.Controls.Add(this.textBoxApi);
            this.panel1.Controls.Add(this.buttonFindFile);
            this.panel1.Controls.Add(this.buttonCategorizeExcel);
            this.panel1.Location = new System.Drawing.Point(0, 2);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(2204, 205);
            this.panel1.TabIndex = 4;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(1841, 7);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(165, 32);
            this.label5.TabIndex = 20;
            this.label5.Text = "Custom fields:";
            this.label5.Click += new System.EventHandler(this.label5_Click);
            // 
            // checkedListBoxCustomFields
            // 
            this.checkedListBoxCustomFields.FormattingEnabled = true;
            this.checkedListBoxCustomFields.Location = new System.Drawing.Point(1841, 43);
            this.checkedListBoxCustomFields.Name = "checkedListBoxCustomFields";
            this.checkedListBoxCustomFields.Size = new System.Drawing.Size(343, 148);
            this.checkedListBoxCustomFields.TabIndex = 19;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(1346, 4);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(103, 32);
            this.label7.TabIndex = 18;
            this.label7.Text = "Options:";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(865, 4);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(102, 32);
            this.label6.TabIndex = 17;
            this.label6.Text = "Context:";
            // 
            // textBoxOptions
            // 
            this.textBoxOptions.Location = new System.Drawing.Point(1346, 39);
            this.textBoxOptions.Multiline = true;
            this.textBoxOptions.Name = "textBoxOptions";
            this.textBoxOptions.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBoxOptions.Size = new System.Drawing.Size(461, 152);
            this.textBoxOptions.TabIndex = 16;
            this.textBoxOptions.Text = "\"culture\":\"en-GB\",\r\n\"includeDetectedCategories\": true,\r\n\"includeCategoryDetails\":" +
    " true,\r\n\"includeMerchantDetails\": true,\r\n\"includeDebugBreakdown\": true,\r\n\"includ" +
    "eCarbonFootprint\": true";
            // 
            // textBoxContext
            // 
            this.textBoxContext.Location = new System.Drawing.Point(865, 39);
            this.textBoxContext.Multiline = true;
            this.textBoxContext.Name = "textBoxContext";
            this.textBoxContext.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBoxContext.Size = new System.Drawing.Size(457, 152);
            this.textBoxContext.TabIndex = 15;
            this.textBoxContext.Text = "\"defaultCountryCode\": \"ES\",\r\n\"defaultCurrency\":\"EUR\",\r\n\"dataSourceProfile\":\"Defau" +
    "ltSpanish\"";
            this.textBoxContext.TextChanged += new System.EventHandler(this.textBoxContext_TextChanged);
            // 
            // progressBarCategorize
            // 
            this.progressBarCategorize.Location = new System.Drawing.Point(199, 77);
            this.progressBarCategorize.Name = "progressBarCategorize";
            this.progressBarCategorize.Size = new System.Drawing.Size(242, 41);
            this.progressBarCategorize.TabIndex = 12;
            // 
            // comboBoxSheet
            // 
            this.comboBoxSheet.FormattingEnabled = true;
            this.comboBoxSheet.Location = new System.Drawing.Point(199, 19);
            this.comboBoxSheet.Name = "comboBoxSheet";
            this.comboBoxSheet.Size = new System.Drawing.Size(242, 40);
            this.comboBoxSheet.TabIndex = 11;
            this.comboBoxSheet.SelectedIndexChanged += new System.EventHandler(this.comboBoxSheet_SelectedIndexChanged);
            this.comboBoxSheet.Format += new System.Windows.Forms.ListControlConvertEventHandler(this.comboBoxSheet_Format);
            this.comboBoxSheet.SelectedValueChanged += new System.EventHandler(this.comboBoxSheet_SelectedValueChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(456, 72);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(146, 32);
            this.label4.TabIndex = 10;
            this.label4.Text = "Client secret";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(456, 13);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(111, 32);
            this.label3.TabIndex = 9;
            this.label3.Text = "Client ID:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(456, 137);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(53, 32);
            this.label2.TabIndex = 8;
            this.label2.Text = "STS";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 140);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(48, 32);
            this.label1.TabIndex = 7;
            this.label1.Text = "API";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // textBoxClientSecret
            // 
            this.textBoxClientSecret.Location = new System.Drawing.Point(628, 72);
            this.textBoxClientSecret.Name = "textBoxClientSecret";
            this.textBoxClientSecret.Size = new System.Drawing.Size(215, 39);
            this.textBoxClientSecret.TabIndex = 6;
            this.textBoxClientSecret.Text = "MenigaDev2021";
            this.textBoxClientSecret.UseSystemPasswordChar = true;
            // 
            // textBoxClientId
            // 
            this.textBoxClientId.Location = new System.Drawing.Point(628, 15);
            this.textBoxClientId.Name = "textBoxClientId";
            this.textBoxClientId.Size = new System.Drawing.Size(215, 39);
            this.textBoxClientId.TabIndex = 5;
            this.textBoxClientId.Text = "int_api_gateway";
            // 
            // textBoxSts
            // 
            this.textBoxSts.Location = new System.Drawing.Point(515, 134);
            this.textBoxSts.Name = "textBoxSts";
            this.textBoxSts.Size = new System.Drawing.Size(334, 39);
            this.textBoxSts.TabIndex = 4;
            this.textBoxSts.Text = "https://identity.meniga.cloud";
            // 
            // textBoxApi
            // 
            this.textBoxApi.Location = new System.Drawing.Point(79, 133);
            this.textBoxApi.Name = "textBoxApi";
            this.textBoxApi.Size = new System.Drawing.Size(362, 39);
            this.textBoxApi.TabIndex = 3;
            this.textBoxApi.Text = "https://api.meniga.cloud";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(13F, 32F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(2196, 1105);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.dataGridViewExcel);
            this.Name = "Form1";
            this.Text = "Meniga Categorize Excel - v1.0";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewExcel)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private Button buttonCategorizeExcel;
        private Button buttonFindFile;
        private DataGridView dataGridViewExcel;
        private Panel panel1;
        private Label label1;
        private TextBox textBoxClientSecret;
        private TextBox textBoxClientId;
        private TextBox textBoxSts;
        private TextBox textBoxApi;
        private Label label4;
        private Label label3;
        private Label label2;
        private ComboBox comboBoxSheet;
        private ProgressBar progressBarCategorize;
        private Label label7;
        private Label label6;
        private TextBox textBoxOptions;
        private TextBox textBoxContext;
        private Label label5;
        private CheckedListBox checkedListBoxCustomFields;
    }
}