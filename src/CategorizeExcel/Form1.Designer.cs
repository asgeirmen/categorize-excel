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
            buttonCategorizeExcel = new Button();
            buttonFindFile = new Button();
            dataGridViewExcel = new DataGridView();
            panel1 = new Panel();
            label9 = new Label();
            comboBoxApiType = new ComboBox();
            textBoxEnrichment = new TextBox();
            label5 = new Label();
            checkedListBoxCustomFields = new CheckedListBox();
            label7 = new Label();
            label6 = new Label();
            textBoxOptions = new TextBox();
            textBoxContext = new TextBox();
            progressBarCategorize = new ProgressBar();
            comboBoxSheet = new ComboBox();
            label4 = new Label();
            label3 = new Label();
            label2 = new Label();
            label1 = new Label();
            textBoxClientSecret = new TextBox();
            textBoxClientId = new TextBox();
            textBoxSts = new TextBox();
            textBoxApi = new TextBox();
            ((System.ComponentModel.ISupportInitialize)dataGridViewExcel).BeginInit();
            panel1.SuspendLayout();
            SuspendLayout();
            // 
            // buttonCategorizeExcel
            // 
            buttonCategorizeExcel.Location = new Point(12, 72);
            buttonCategorizeExcel.Name = "buttonCategorizeExcel";
            buttonCategorizeExcel.Size = new Size(150, 46);
            buttonCategorizeExcel.TabIndex = 1;
            buttonCategorizeExcel.Text = "Categorize";
            buttonCategorizeExcel.UseVisualStyleBackColor = true;
            buttonCategorizeExcel.Click += buttonCategorizeExcel_Click;
            // 
            // buttonFindFile
            // 
            buttonFindFile.Location = new Point(12, 16);
            buttonFindFile.Name = "buttonFindFile";
            buttonFindFile.Size = new Size(150, 46);
            buttonFindFile.TabIndex = 2;
            buttonFindFile.Text = "Load excel...";
            buttonFindFile.UseVisualStyleBackColor = true;
            buttonFindFile.Click += buttonFindFile_Click;
            // 
            // dataGridViewExcel
            // 
            dataGridViewExcel.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            dataGridViewExcel.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewExcel.Location = new Point(0, 304);
            dataGridViewExcel.Name = "dataGridViewExcel";
            dataGridViewExcel.RowHeadersWidth = 82;
            dataGridViewExcel.RowTemplate.Height = 41;
            dataGridViewExcel.Size = new Size(2202, 797);
            dataGridViewExcel.TabIndex = 3;
            // 
            // panel1
            // 
            panel1.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            panel1.Controls.Add(label9);
            panel1.Controls.Add(comboBoxApiType);
            panel1.Controls.Add(textBoxEnrichment);
            panel1.Controls.Add(label5);
            panel1.Controls.Add(checkedListBoxCustomFields);
            panel1.Controls.Add(label7);
            panel1.Controls.Add(label6);
            panel1.Controls.Add(textBoxOptions);
            panel1.Controls.Add(textBoxContext);
            panel1.Controls.Add(progressBarCategorize);
            panel1.Controls.Add(comboBoxSheet);
            panel1.Controls.Add(label4);
            panel1.Controls.Add(label3);
            panel1.Controls.Add(label2);
            panel1.Controls.Add(label1);
            panel1.Controls.Add(textBoxClientSecret);
            panel1.Controls.Add(textBoxClientId);
            panel1.Controls.Add(textBoxSts);
            panel1.Controls.Add(textBoxApi);
            panel1.Controls.Add(buttonFindFile);
            panel1.Controls.Add(buttonCategorizeExcel);
            panel1.Location = new Point(0, 2);
            panel1.Name = "panel1";
            panel1.Size = new Size(2204, 249);
            panel1.TabIndex = 4;
            // 
            // label9
            // 
            label9.AutoSize = true;
            label9.Location = new Point(12, 195);
            label9.Name = "label9";
            label9.Size = new Size(102, 32);
            label9.TabIndex = 24;
            label9.Text = "API type";
            // 
            // comboBoxApiType
            // 
            comboBoxApiType.FormattingEnabled = true;
            comboBoxApiType.Items.AddRange(new object[] { "Core", "Enrichment", "TapiX", "Snowdrop" });
            comboBoxApiType.Location = new Point(168, 187);
            comboBoxApiType.Name = "comboBoxApiType";
            comboBoxApiType.Size = new Size(273, 40);
            comboBoxApiType.TabIndex = 22;
            comboBoxApiType.Text = "Core";
            // 
            // textBoxEnrichment
            // 
            textBoxEnrichment.Enabled = false;
            textBoxEnrichment.Location = new Point(456, 192);
            textBoxEnrichment.Name = "textBoxEnrichment";
            textBoxEnrichment.Size = new Size(387, 39);
            textBoxEnrichment.TabIndex = 21;
            textBoxEnrichment.Text = "http://localhost:20052";
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.Location = new Point(1841, 7);
            label5.Name = "label5";
            label5.Size = new Size(165, 32);
            label5.TabIndex = 20;
            label5.Text = "Custom fields:";
            // 
            // checkedListBoxCustomFields
            // 
            checkedListBoxCustomFields.FormattingEnabled = true;
            checkedListBoxCustomFields.Location = new Point(1841, 43);
            checkedListBoxCustomFields.Name = "checkedListBoxCustomFields";
            checkedListBoxCustomFields.Size = new Size(343, 148);
            checkedListBoxCustomFields.TabIndex = 19;
            // 
            // label7
            // 
            label7.AutoSize = true;
            label7.Location = new Point(1346, 4);
            label7.Name = "label7";
            label7.Size = new Size(103, 32);
            label7.TabIndex = 18;
            label7.Text = "Options:";
            // 
            // label6
            // 
            label6.AutoSize = true;
            label6.Location = new Point(865, 4);
            label6.Name = "label6";
            label6.Size = new Size(102, 32);
            label6.TabIndex = 17;
            label6.Text = "Context:";
            // 
            // textBoxOptions
            // 
            textBoxOptions.Location = new Point(1346, 39);
            textBoxOptions.Multiline = true;
            textBoxOptions.Name = "textBoxOptions";
            textBoxOptions.ScrollBars = ScrollBars.Both;
            textBoxOptions.Size = new Size(461, 152);
            textBoxOptions.TabIndex = 16;
            textBoxOptions.Text = "\"culture\":\"en-GB\",\r\n\"includeDetectedCategories\": true,\r\n\"includeCategoryDetails\": true,\r\n\"includeMerchantDetails\": true,\r\n\"includeDebugBreakdown\": true,\r\n\"includeCarbonFootprint\": false";
            // 
            // textBoxContext
            // 
            textBoxContext.Location = new Point(865, 39);
            textBoxContext.Multiline = true;
            textBoxContext.Name = "textBoxContext";
            textBoxContext.ScrollBars = ScrollBars.Both;
            textBoxContext.Size = new Size(457, 152);
            textBoxContext.TabIndex = 15;
            textBoxContext.Text = "\"defaultCountryCode\": \"ES\",\r\n\"defaultCurrency\":\"EUR\"";
            // 
            // progressBarCategorize
            // 
            progressBarCategorize.Location = new Point(168, 77);
            progressBarCategorize.Name = "progressBarCategorize";
            progressBarCategorize.Size = new Size(273, 41);
            progressBarCategorize.TabIndex = 12;
            // 
            // comboBoxSheet
            // 
            comboBoxSheet.FormattingEnabled = true;
            comboBoxSheet.Location = new Point(168, 19);
            comboBoxSheet.Name = "comboBoxSheet";
            comboBoxSheet.Size = new Size(273, 40);
            comboBoxSheet.TabIndex = 11;
            comboBoxSheet.SelectedIndexChanged += comboBoxSheet_SelectedIndexChanged;
            comboBoxSheet.Format += comboBoxSheet_Format;
            comboBoxSheet.SelectedValueChanged += comboBoxSheet_SelectedValueChanged;
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Location = new Point(456, 72);
            label4.Name = "label4";
            label4.Size = new Size(146, 32);
            label4.TabIndex = 10;
            label4.Text = "Client secret";
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(456, 13);
            label3.Name = "label3";
            label3.Size = new Size(111, 32);
            label3.TabIndex = 9;
            label3.Text = "Client ID:";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(456, 137);
            label2.Name = "label2";
            label2.Size = new Size(53, 32);
            label2.TabIndex = 8;
            label2.Text = "STS";
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(12, 140);
            label1.Name = "label1";
            label1.Size = new Size(48, 32);
            label1.TabIndex = 7;
            label1.Text = "API";
            label1.Click += label1_Click;
            // 
            // textBoxClientSecret
            // 
            textBoxClientSecret.Location = new Point(628, 72);
            textBoxClientSecret.Name = "textBoxClientSecret";
            textBoxClientSecret.Size = new Size(215, 39);
            textBoxClientSecret.TabIndex = 6;
            textBoxClientSecret.Text = "rootAdmin321";
            textBoxClientSecret.UseSystemPasswordChar = true;
            // 
            // textBoxClientId
            // 
            textBoxClientId.Location = new Point(628, 15);
            textBoxClientId.Name = "textBoxClientId";
            textBoxClientId.Size = new Size(215, 39);
            textBoxClientId.TabIndex = 5;
            textBoxClientId.Text = "int_api_gateway";
            // 
            // textBoxSts
            // 
            textBoxSts.Location = new Point(515, 134);
            textBoxSts.Name = "textBoxSts";
            textBoxSts.Size = new Size(334, 39);
            textBoxSts.TabIndex = 4;
            textBoxSts.Text = "https://localhost:20011";
            // 
            // textBoxApi
            // 
            textBoxApi.Location = new Point(168, 133);
            textBoxApi.Name = "textBoxApi";
            textBoxApi.Size = new Size(273, 39);
            textBoxApi.TabIndex = 3;
            textBoxApi.Text = "http://localhost:20000";
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(13F, 32F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(2196, 1105);
            Controls.Add(panel1);
            Controls.Add(dataGridViewExcel);
            Name = "Form1";
            Text = "Meniga Categorize Excel - v1.2";
            ((System.ComponentModel.ISupportInitialize)dataGridViewExcel).EndInit();
            panel1.ResumeLayout(false);
            panel1.PerformLayout();
            ResumeLayout(false);
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
        private TextBox textBoxEnrichment;
        private Label label9;
        private ComboBox comboBoxApiType;
    }
}