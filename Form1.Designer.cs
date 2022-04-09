namespace ExcelFileModify
{
    partial class Form1
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
            this.importData = new System.Windows.Forms.Button();
            this.modify_excel_file = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.clearData = new System.Windows.Forms.Button();
            this.inParts_chb = new System.Windows.Forms.CheckBox();
            this.statusBar = new System.Windows.Forms.StatusBar();
            this.SuspendLayout();
            // 
            // importData
            // 
            this.importData.Cursor = System.Windows.Forms.Cursors.Hand;
            this.importData.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.importData.Location = new System.Drawing.Point(58, 114);
            this.importData.Name = "importData";
            this.importData.Size = new System.Drawing.Size(160, 71);
            this.importData.TabIndex = 0;
            this.importData.Text = "&Import";
            this.importData.UseVisualStyleBackColor = true;
            this.importData.Click += new System.EventHandler(this.importData_click);
            // 
            // modify_excel_file
            // 
            this.modify_excel_file.Cursor = System.Windows.Forms.Cursors.Hand;
            this.modify_excel_file.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.modify_excel_file.Location = new System.Drawing.Point(275, 114);
            this.modify_excel_file.Name = "modify_excel_file";
            this.modify_excel_file.Size = new System.Drawing.Size(88, 71);
            this.modify_excel_file.TabIndex = 1;
            this.modify_excel_file.Text = "&Add Zones";
            this.modify_excel_file.UseVisualStyleBackColor = true;
            this.modify_excel_file.Click += new System.EventHandler(this.modify_excel_file_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(58, 369);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(100, 23);
            this.progressBar1.TabIndex = 2;
            this.progressBar1.Visible = false;
            // 
            // textBox1
            // 
            this.textBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox1.Location = new System.Drawing.Point(58, 245);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.Size = new System.Drawing.Size(150, 98);
            this.textBox1.TabIndex = 5;
            this.textBox1.Text = "1 toxum улица, зоны доставки";
            // 
            // textBox2
            // 
            this.textBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox2.Location = new System.Drawing.Point(264, 245);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(177, 30);
            this.textBox2.TabIndex = 6;
            this.textBox2.Text = "1 toxum улица";
            // 
            // clearData
            // 
            this.clearData.Cursor = System.Windows.Forms.Cursors.Hand;
            this.clearData.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.clearData.Location = new System.Drawing.Point(415, 114);
            this.clearData.Name = "clearData";
            this.clearData.Size = new System.Drawing.Size(93, 71);
            this.clearData.TabIndex = 7;
            this.clearData.Text = "&Delete";
            this.clearData.UseVisualStyleBackColor = true;
            this.clearData.Click += new System.EventHandler(this.clearData_Click);
            // 
            // inParts_chb
            // 
            this.inParts_chb.AutoSize = true;
            this.inParts_chb.Checked = true;
            this.inParts_chb.CheckState = System.Windows.Forms.CheckState.Checked;
            this.inParts_chb.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.inParts_chb.Location = new System.Drawing.Point(475, 32);
            this.inParts_chb.Name = "inParts_chb";
            this.inParts_chb.Size = new System.Drawing.Size(87, 24);
            this.inParts_chb.TabIndex = 8;
            this.inParts_chb.Text = "in parts";
            this.inParts_chb.UseVisualStyleBackColor = true;
            // 
            // statusBar
            // 
            this.statusBar.Font = new System.Drawing.Font("Verdana", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.statusBar.Location = new System.Drawing.Point(0, 408);
            this.statusBar.Name = "statusBar";
            this.statusBar.Size = new System.Drawing.Size(582, 42);
            this.statusBar.TabIndex = 9;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(582, 450);
            this.Controls.Add(this.statusBar);
            this.Controls.Add(this.inParts_chb);
            this.Controls.Add(this.clearData);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.modify_excel_file);
            this.Controls.Add(this.importData);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "Form1";
            this.Text = "Excel";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button importData;
        private System.Windows.Forms.Button modify_excel_file;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Button clearData;
        private System.Windows.Forms.CheckBox inParts_chb;
        private System.Windows.Forms.StatusBar statusBar;
    }
}

