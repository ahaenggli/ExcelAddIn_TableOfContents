namespace ExcelAddIn_TableOfContents
{
    partial class frmTocSheetExtension
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.cbSetDefault = new System.Windows.Forms.CheckBox();
            this.OK = new System.Windows.Forms.Button();
            this.txtSumTitel = new System.Windows.Forms.TextBox();
            this.txtWorkSheetCreatedDate = new System.Windows.Forms.TextBox();
            this.txtProperties = new System.Windows.Forms.TextBox();
            this.txtSummaryColumns = new System.Windows.Forms.TextBox();
            this.txtStyle = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(108, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Name Übersichtsblatt";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(155, 9);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(110, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Name für Erstelldatum";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 58);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(258, 13);
            this.label3.TabIndex = 2;
            this.label3.Text = "Felder für die einzelnen Sheets: (Trennzeichen ist \";\")";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(12, 145);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(257, 13);
            this.label4.TabIndex = 3;
            this.label4.Text = "Spalten für das Übersichtsblatt: (Trennzeichen ist \";\")";
            // 
            // cbSetDefault
            // 
            this.cbSetDefault.AutoSize = true;
            this.cbSetDefault.Location = new System.Drawing.Point(15, 233);
            this.cbSetDefault.Name = "cbSetDefault";
            this.cbSetDefault.Size = new System.Drawing.Size(184, 17);
            this.cbSetDefault.TabIndex = 4;
            this.cbSetDefault.Text = "Standard für neue Arbeitsmappen";
            this.cbSetDefault.UseVisualStyleBackColor = true;
            // 
            // OK
            // 
            this.OK.Location = new System.Drawing.Point(345, 227);
            this.OK.Name = "OK";
            this.OK.Size = new System.Drawing.Size(75, 23);
            this.OK.TabIndex = 6;
            this.OK.Text = "OK";
            this.OK.UseVisualStyleBackColor = true;
            this.OK.Click += new System.EventHandler(this.OK_Click);
            // 
            // txtSumTitel
            // 
            this.txtSumTitel.Location = new System.Drawing.Point(15, 25);
            this.txtSumTitel.Name = "txtSumTitel";
            this.txtSumTitel.Size = new System.Drawing.Size(100, 20);
            this.txtSumTitel.TabIndex = 7;
            // 
            // txtWorkSheetCreatedDate
            // 
            this.txtWorkSheetCreatedDate.Location = new System.Drawing.Point(158, 25);
            this.txtWorkSheetCreatedDate.Name = "txtWorkSheetCreatedDate";
            this.txtWorkSheetCreatedDate.Size = new System.Drawing.Size(100, 20);
            this.txtWorkSheetCreatedDate.TabIndex = 8;
            // 
            // txtProperties
            // 
            this.txtProperties.Location = new System.Drawing.Point(15, 74);
            this.txtProperties.Multiline = true;
            this.txtProperties.Name = "txtProperties";
            this.txtProperties.Size = new System.Drawing.Size(405, 55);
            this.txtProperties.TabIndex = 9;
            // 
            // txtSummaryColumns
            // 
            this.txtSummaryColumns.Location = new System.Drawing.Point(15, 161);
            this.txtSummaryColumns.Multiline = true;
            this.txtSummaryColumns.Name = "txtSummaryColumns";
            this.txtSummaryColumns.Size = new System.Drawing.Size(405, 52);
            this.txtSummaryColumns.TabIndex = 10;
            // 
            // txtStyle
            // 
            this.txtStyle.Location = new System.Drawing.Point(285, 25);
            this.txtStyle.Name = "txtStyle";
            this.txtStyle.Size = new System.Drawing.Size(135, 20);
            this.txtStyle.TabIndex = 12;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(282, 9);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(103, 13);
            this.label5.TabIndex = 11;
            this.label5.Text = "Style Übersichtsblatt";
            // 
            // TocSheetExtensionForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(434, 271);
            this.Controls.Add(this.txtStyle);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.txtSummaryColumns);
            this.Controls.Add(this.txtProperties);
            this.Controls.Add(this.txtWorkSheetCreatedDate);
            this.Controls.Add(this.txtSumTitel);
            this.Controls.Add(this.OK);
            this.Controls.Add(this.cbSetDefault);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Name = "TocSheetExtensionForm";
            this.Text = "edit custom values for index sheet";
            this.Load += new System.EventHandler(this.TocSheetExtensionForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.CheckBox cbSetDefault;
        private System.Windows.Forms.Button OK;
        private System.Windows.Forms.TextBox txtSumTitel;
        private System.Windows.Forms.TextBox txtWorkSheetCreatedDate;
        private System.Windows.Forms.TextBox txtProperties;
        private System.Windows.Forms.TextBox txtSummaryColumns;
        private System.Windows.Forms.TextBox txtStyle;
        private System.Windows.Forms.Label label5;
    }
}