namespace ExcelParser
{
    partial class Form_ExcelParser
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
            TB_ExcelList = new TextBox();
            label1 = new Label();
            label2 = new Label();
            TB_CsPath = new TextBox();
            label3 = new Label();
            TB_JsonPath = new TextBox();
            Btn_LoadExcelListPath = new Button();
            Btn_SaveCsPath = new Button();
            Btn_SaveJsonPath = new Button();
            Btn_ConvertAll = new Button();
            Btn_ConvertToJson = new Button();
            Btn_ConvertToCs = new Button();
            Label_Desc = new Label();
            Btn_LoadExcelPath = new Button();
            label5 = new Label();
            TB_ExcelPath = new TextBox();
            SuspendLayout();
            // 
            // TB_ExcelList
            // 
            TB_ExcelList.Location = new Point(41, 57);
            TB_ExcelList.Name = "TB_ExcelList";
            TB_ExcelList.Size = new Size(579, 31);
            TB_ExcelList.TabIndex = 0;
            TB_ExcelList.TabStop = false;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(41, 29);
            label1.Name = "label1";
            label1.Size = new Size(252, 25);
            label1.TabIndex = 1;
            label1.Text = "변환할 엑셀 파일 리스트 경로";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(41, 179);
            label2.Name = "label2";
            label2.Size = new Size(155, 25);
            label2.TabIndex = 3;
            label2.Text = "cs 파일 저장 경로";
            // 
            // TB_CsPath
            // 
            TB_CsPath.Location = new Point(41, 207);
            TB_CsPath.Name = "TB_CsPath";
            TB_CsPath.Size = new Size(579, 31);
            TB_CsPath.TabIndex = 2;
            TB_CsPath.TabStop = false;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(41, 258);
            label3.Name = "label3";
            label3.Size = new Size(171, 25);
            label3.TabIndex = 5;
            label3.Text = "json 파일 저장 경로";
            // 
            // TB_JsonPath
            // 
            TB_JsonPath.Location = new Point(41, 286);
            TB_JsonPath.Name = "TB_JsonPath";
            TB_JsonPath.Size = new Size(579, 31);
            TB_JsonPath.TabIndex = 4;
            TB_JsonPath.TabStop = false;
            // 
            // Btn_LoadExcelListPath
            // 
            Btn_LoadExcelListPath.Location = new Point(655, 51);
            Btn_LoadExcelListPath.Name = "Btn_LoadExcelListPath";
            Btn_LoadExcelListPath.Size = new Size(93, 42);
            Btn_LoadExcelListPath.TabIndex = 6;
            Btn_LoadExcelListPath.Text = "불러오기";
            Btn_LoadExcelListPath.UseVisualStyleBackColor = true;
            Btn_LoadExcelListPath.Click += Btn_LoadExcelListPath_Click;
            // 
            // Btn_SaveCsPath
            // 
            Btn_SaveCsPath.Location = new Point(655, 201);
            Btn_SaveCsPath.Name = "Btn_SaveCsPath";
            Btn_SaveCsPath.Size = new Size(93, 42);
            Btn_SaveCsPath.TabIndex = 7;
            Btn_SaveCsPath.Text = "불러오기";
            Btn_SaveCsPath.UseVisualStyleBackColor = true;
            Btn_SaveCsPath.Click += Btn_SaveCsPath_Click;
            // 
            // Btn_SaveJsonPath
            // 
            Btn_SaveJsonPath.Location = new Point(655, 280);
            Btn_SaveJsonPath.Name = "Btn_SaveJsonPath";
            Btn_SaveJsonPath.Size = new Size(93, 42);
            Btn_SaveJsonPath.TabIndex = 8;
            Btn_SaveJsonPath.Text = "불러오기";
            Btn_SaveJsonPath.UseVisualStyleBackColor = true;
            Btn_SaveJsonPath.Click += Btn_SaveJsonPath_Click;
            // 
            // Btn_ConvertAll
            // 
            Btn_ConvertAll.Location = new Point(655, 383);
            Btn_ConvertAll.Name = "Btn_ConvertAll";
            Btn_ConvertAll.Size = new Size(111, 42);
            Btn_ConvertAll.TabIndex = 9;
            Btn_ConvertAll.Text = "모두 변환";
            Btn_ConvertAll.UseVisualStyleBackColor = true;
            Btn_ConvertAll.Click += Btn_ConvertAll_Click;
            // 
            // Btn_ConvertToJson
            // 
            Btn_ConvertToJson.Location = new Point(479, 383);
            Btn_ConvertToJson.Name = "Btn_ConvertToJson";
            Btn_ConvertToJson.Size = new Size(155, 42);
            Btn_ConvertToJson.TabIndex = 10;
            Btn_ConvertToJson.Text = "json 파일로 변환";
            Btn_ConvertToJson.UseVisualStyleBackColor = true;
            Btn_ConvertToJson.Click += Btn_ConvertToJson_Click;
            // 
            // Btn_ConvertToCs
            // 
            Btn_ConvertToCs.Location = new Point(303, 383);
            Btn_ConvertToCs.Name = "Btn_ConvertToCs";
            Btn_ConvertToCs.Size = new Size(155, 42);
            Btn_ConvertToCs.TabIndex = 11;
            Btn_ConvertToCs.Text = "cs 파일로 변환";
            Btn_ConvertToCs.UseVisualStyleBackColor = true;
            Btn_ConvertToCs.Click += Btn_ConvertToCs_Click;
            // 
            // Label_Desc
            // 
            Label_Desc.AutoSize = true;
            Label_Desc.Location = new Point(41, 340);
            Label_Desc.Name = "Label_Desc";
            Label_Desc.Size = new Size(24, 25);
            Label_Desc.TabIndex = 12;
            Label_Desc.Text = "...";
            Label_Desc.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // Btn_LoadExcelPath
            // 
            Btn_LoadExcelPath.Location = new Point(655, 124);
            Btn_LoadExcelPath.Name = "Btn_LoadExcelPath";
            Btn_LoadExcelPath.Size = new Size(93, 42);
            Btn_LoadExcelPath.TabIndex = 15;
            Btn_LoadExcelPath.Text = "불러오기";
            Btn_LoadExcelPath.UseVisualStyleBackColor = true;
            Btn_LoadExcelPath.Click += Btn_LoadExcelPath_Click;
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.Location = new Point(41, 102);
            label5.Name = "label5";
            label5.Size = new Size(132, 25);
            label5.TabIndex = 14;
            label5.Text = "엑셀 파일 경로";
            // 
            // TB_ExcelPath
            // 
            TB_ExcelPath.Location = new Point(41, 130);
            TB_ExcelPath.Name = "TB_ExcelPath";
            TB_ExcelPath.Size = new Size(579, 31);
            TB_ExcelPath.TabIndex = 13;
            TB_ExcelPath.TabStop = false;
            // 
            // Form_ExcelParser
            // 
            AutoScaleDimensions = new SizeF(10F, 25F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(789, 437);
            Controls.Add(Btn_LoadExcelPath);
            Controls.Add(label5);
            Controls.Add(TB_ExcelPath);
            Controls.Add(Label_Desc);
            Controls.Add(Btn_ConvertToCs);
            Controls.Add(Btn_ConvertToJson);
            Controls.Add(Btn_ConvertAll);
            Controls.Add(Btn_SaveJsonPath);
            Controls.Add(Btn_SaveCsPath);
            Controls.Add(Btn_LoadExcelListPath);
            Controls.Add(label3);
            Controls.Add(TB_JsonPath);
            Controls.Add(label2);
            Controls.Add(TB_CsPath);
            Controls.Add(label1);
            Controls.Add(TB_ExcelList);
            Name = "Form_ExcelParser";
            Text = "엑셀 변환기";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private TextBox TB_ExcelList;
        private Label label1;
        private Label label2;
        private TextBox TB_CsPath;
        private Label label3;
        private TextBox TB_JsonPath;
        private Button Btn_LoadExcelListPath;
        private Button Btn_SaveCsPath;
        private Button Btn_SaveJsonPath;
        private Button Btn_ConvertAll;
        private Button Btn_ConvertToJson;
        private Button Btn_ConvertToCs;
        private Label Label_Desc;
        private Button Btn_LoadExcelPath;
        private Label label5;
        private TextBox TB_ExcelPath;
    }
}
