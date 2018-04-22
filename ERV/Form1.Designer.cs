namespace ERV
{
    partial class mainForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(mainForm));
            this.dgResults = new System.Windows.Forms.DataGridView();
            this.cbUsers = new System.Windows.Forms.ComboBox();
            this.btnCalculate = new System.Windows.Forms.Button();
            this.comboBox2 = new System.Windows.Forms.ComboBox();
            this.lblUser = new System.Windows.Forms.Label();
            this.lblMonth = new System.Windows.Forms.Label();
            this.lblTotalTime = new System.Windows.Forms.Label();
            this.totalTime = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.tbLocation = new System.Windows.Forms.TextBox();
            this.btnPrint = new System.Windows.Forms.Button();
            this.dgPrintDocument = new System.Drawing.Printing.PrintDocument();
            ((System.ComponentModel.ISupportInitialize)(this.dgResults)).BeginInit();
            this.SuspendLayout();
            // 
            // dgResults
            // 
            this.dgResults.AllowUserToAddRows = false;
            this.dgResults.AllowUserToDeleteRows = false;
            this.dgResults.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgResults.Location = new System.Drawing.Point(12, 63);
            this.dgResults.Name = "dgResults";
            this.dgResults.ReadOnly = true;
            this.dgResults.Size = new System.Drawing.Size(776, 375);
            this.dgResults.TabIndex = 0;
            // 
            // cbUsers
            // 
            this.cbUsers.DisplayMember = "USERID";
            this.cbUsers.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbUsers.FormattingEnabled = true;
            this.cbUsers.Location = new System.Drawing.Point(12, 29);
            this.cbUsers.Name = "cbUsers";
            this.cbUsers.Size = new System.Drawing.Size(121, 28);
            this.cbUsers.TabIndex = 1;
            this.cbUsers.ValueMember = "USERID";
            // 
            // btnCalculate
            // 
            this.btnCalculate.Location = new System.Drawing.Point(386, 13);
            this.btnCalculate.Name = "btnCalculate";
            this.btnCalculate.Size = new System.Drawing.Size(114, 44);
            this.btnCalculate.TabIndex = 2;
            this.btnCalculate.Text = "Izračunaj";
            this.btnCalculate.UseVisualStyleBackColor = true;
            this.btnCalculate.Click += new System.EventHandler(this.button1_Click);
            // 
            // comboBox2
            // 
            this.comboBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comboBox2.FormattingEnabled = true;
            this.comboBox2.Location = new System.Drawing.Point(139, 29);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(121, 28);
            this.comboBox2.TabIndex = 3;
            // 
            // lblUser
            // 
            this.lblUser.AutoSize = true;
            this.lblUser.Location = new System.Drawing.Point(9, 13);
            this.lblUser.Name = "lblUser";
            this.lblUser.Size = new System.Drawing.Size(41, 13);
            this.lblUser.TabIndex = 4;
            this.lblUser.Text = "Radnik";
            // 
            // lblMonth
            // 
            this.lblMonth.AutoSize = true;
            this.lblMonth.Location = new System.Drawing.Point(136, 13);
            this.lblMonth.Name = "lblMonth";
            this.lblMonth.Size = new System.Drawing.Size(41, 13);
            this.lblMonth.TabIndex = 5;
            this.lblMonth.Text = "Mjesec";
            // 
            // lblTotalTime
            // 
            this.lblTotalTime.AutoSize = true;
            this.lblTotalTime.Location = new System.Drawing.Point(652, 9);
            this.lblTotalTime.Name = "lblTotalTime";
            this.lblTotalTime.Size = new System.Drawing.Size(136, 13);
            this.lblTotalTime.TabIndex = 6;
            this.lblTotalTime.Text = "Ukupno radno vrijeme (sati)";
            this.lblTotalTime.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // totalTime
            // 
            this.totalTime.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.totalTime.Location = new System.Drawing.Point(677, 29);
            this.totalTime.Name = "totalTime";
            this.totalTime.Size = new System.Drawing.Size(111, 28);
            this.totalTime.TabIndex = 7;
            this.totalTime.Text = "0";
            this.totalTime.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 452);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(108, 13);
            this.label1.TabIndex = 8;
            this.label1.Text = "Naziv baze podataka";
            // 
            // tbLocation
            // 
            this.tbLocation.Location = new System.Drawing.Point(136, 449);
            this.tbLocation.Name = "tbLocation";
            this.tbLocation.ReadOnly = true;
            this.tbLocation.Size = new System.Drawing.Size(652, 20);
            this.tbLocation.TabIndex = 9;
            // 
            // btnPrint
            // 
            this.btnPrint.Location = new System.Drawing.Point(266, 13);
            this.btnPrint.Name = "btnPrint";
            this.btnPrint.Size = new System.Drawing.Size(114, 44);
            this.btnPrint.TabIndex = 10;
            this.btnPrint.Text = "Print";
            this.btnPrint.UseVisualStyleBackColor = true;
            this.btnPrint.Click += new System.EventHandler(this.btnPrint_Click);
            // 
            // dgPrintDocument
            // 
            this.dgPrintDocument.BeginPrint += new System.Drawing.Printing.PrintEventHandler(this.dgPrintDocument_BeginPrint);
            this.dgPrintDocument.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(this.dgPrintDocument_PrintPage);
            // 
            // mainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 477);
            this.Controls.Add(this.btnPrint);
            this.Controls.Add(this.tbLocation);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.totalTime);
            this.Controls.Add(this.lblTotalTime);
            this.Controls.Add(this.lblMonth);
            this.Controls.Add(this.lblUser);
            this.Controls.Add(this.comboBox2);
            this.Controls.Add(this.btnCalculate);
            this.Controls.Add(this.cbUsers);
            this.Controls.Add(this.dgResults);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "mainForm";
            this.Text = "Evidencija Radnog Vremena";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgResults)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dgResults;
        private System.Windows.Forms.ComboBox cbUsers;
        private System.Windows.Forms.Button btnCalculate;
        private System.Windows.Forms.ComboBox comboBox2;
        private System.Windows.Forms.Label lblUser;
        private System.Windows.Forms.Label lblMonth;
        private System.Windows.Forms.Label lblTotalTime;
        private System.Windows.Forms.Label totalTime;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tbLocation;
        private System.Windows.Forms.Button btnPrint;
        private System.Drawing.Printing.PrintDocument dgPrintDocument;
    }
}

