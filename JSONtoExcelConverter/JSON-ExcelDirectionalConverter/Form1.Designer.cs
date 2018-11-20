namespace JSON_ExcelDirectionalConverter
{
    partial class Form1
    {
        /// <summary>
        /// 필수 디자이너 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 사용 중인 모든 리소스를 정리합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 디자이너에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다.
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마십시오.
        /// </summary>
        private void InitializeComponent()
        {
            this.btn_convert = new System.Windows.Forms.Button();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.lb_fileList = new System.Windows.Forms.ListBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btn_addFiles = new System.Windows.Forms.Button();
            this.btn_removeFiles = new System.Windows.Forms.Button();
            this.btn_clearList = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.cb_modeSelect = new System.Windows.Forms.ComboBox();
            this.groupBox3.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btn_convert
            // 
            this.btn_convert.Font = new System.Drawing.Font("굴림", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.btn_convert.Location = new System.Drawing.Point(242, 258);
            this.btn_convert.Name = "btn_convert";
            this.btn_convert.Size = new System.Drawing.Size(171, 28);
            this.btn_convert.TabIndex = 15;
            this.btn_convert.Text = "변환";
            this.btn_convert.UseVisualStyleBackColor = true;
            this.btn_convert.Click += new System.EventHandler(this.btn_convert_Click);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.lb_fileList);
            this.groupBox3.Location = new System.Drawing.Point(20, 24);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(200, 261);
            this.groupBox3.TabIndex = 14;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "파일 목록";
            // 
            // lb_fileList
            // 
            this.lb_fileList.FormattingEnabled = true;
            this.lb_fileList.HorizontalScrollbar = true;
            this.lb_fileList.ItemHeight = 12;
            this.lb_fileList.Location = new System.Drawing.Point(6, 20);
            this.lb_fileList.Name = "lb_fileList";
            this.lb_fileList.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.lb_fileList.Size = new System.Drawing.Size(186, 232);
            this.lb_fileList.TabIndex = 0;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.btn_addFiles);
            this.groupBox2.Controls.Add(this.btn_removeFiles);
            this.groupBox2.Controls.Add(this.btn_clearList);
            this.groupBox2.Location = new System.Drawing.Point(252, 97);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(161, 155);
            this.groupBox2.TabIndex = 13;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "목록 관리";
            // 
            // btn_addFiles
            // 
            this.btn_addFiles.Font = new System.Drawing.Font("굴림", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.btn_addFiles.Location = new System.Drawing.Point(20, 20);
            this.btn_addFiles.Name = "btn_addFiles";
            this.btn_addFiles.Size = new System.Drawing.Size(120, 23);
            this.btn_addFiles.TabIndex = 2;
            this.btn_addFiles.Text = "파일 추가";
            this.btn_addFiles.UseVisualStyleBackColor = true;
            this.btn_addFiles.Click += new System.EventHandler(this.btn_addFiles_Click);
            // 
            // btn_removeFiles
            // 
            this.btn_removeFiles.Font = new System.Drawing.Font("굴림", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.btn_removeFiles.Location = new System.Drawing.Point(20, 66);
            this.btn_removeFiles.Name = "btn_removeFiles";
            this.btn_removeFiles.Size = new System.Drawing.Size(120, 23);
            this.btn_removeFiles.TabIndex = 3;
            this.btn_removeFiles.Text = "파일 삭제";
            this.btn_removeFiles.UseVisualStyleBackColor = true;
            this.btn_removeFiles.Click += new System.EventHandler(this.btn_removeFiles_Click);
            // 
            // btn_clearList
            // 
            this.btn_clearList.Font = new System.Drawing.Font("굴림", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.btn_clearList.Location = new System.Drawing.Point(20, 112);
            this.btn_clearList.Name = "btn_clearList";
            this.btn_clearList.Size = new System.Drawing.Size(120, 23);
            this.btn_clearList.TabIndex = 5;
            this.btn_clearList.Text = "목록 초기화";
            this.btn_clearList.UseVisualStyleBackColor = true;
            this.btn_clearList.Click += new System.EventHandler(this.btn_clearList_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.cb_modeSelect);
            this.groupBox1.Location = new System.Drawing.Point(252, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(161, 79);
            this.groupBox1.TabIndex = 12;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "모드 선택";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("굴림", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.label1.Location = new System.Drawing.Point(13, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(133, 15);
            this.label1.TabIndex = 6;
            this.label1.Text = "-변환 모드 선택-";
            // 
            // cb_modeSelect
            // 
            this.cb_modeSelect.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cb_modeSelect.Font = new System.Drawing.Font("굴림", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.cb_modeSelect.FormattingEnabled = true;
            this.cb_modeSelect.Items.AddRange(new object[] {
            "WJSON to WExcel",
            "WExcel to WJSON",
            "CJSON to CExcel",
            "CExcel to EJSON"});
            this.cb_modeSelect.Location = new System.Drawing.Point(19, 42);
            this.cb_modeSelect.Name = "cb_modeSelect";
            this.cb_modeSelect.Size = new System.Drawing.Size(121, 23);
            this.cb_modeSelect.TabIndex = 1;
            this.cb_modeSelect.SelectedIndexChanged += new System.EventHandler(this.cb_modeSelect_SelectedIndexChanged);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(436, 298);
            this.Controls.Add(this.btn_convert);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.groupBox3.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btn_convert;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.ListBox lb_fileList;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button btn_addFiles;
        private System.Windows.Forms.Button btn_removeFiles;
        private System.Windows.Forms.Button btn_clearList;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cb_modeSelect;
    }
}

