namespace LatihanADONET
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
            this.btnTesKoneksi = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.lvwMahasiswa = new System.Windows.Forms.ListView();
            this.btnTampilkanData = new System.Windows.Forms.Button();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.btnInsert = new System.Windows.Forms.Button();
            this.txtAngkatanInsert = new System.Windows.Forms.TextBox();
            this.txtNamaInsert = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtNpmInsert = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.btnCariUpdate = new System.Windows.Forms.Button();
            this.btnUpdate = new System.Windows.Forms.Button();
            this.txtAngkatanUpdate = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtNamaUpdate = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.txtNpmUpdate = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.tabPage4 = new System.Windows.Forms.TabPage();
            this.btnCariDelete = new System.Windows.Forms.Button();
            this.btnDelete = new System.Windows.Forms.Button();
            this.label9 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.txtAngkatanDelete = new System.Windows.Forms.TextBox();
            this.txtNamaDelete = new System.Windows.Forms.TextBox();
            this.txtNpmDelete = new System.Windows.Forms.TextBox();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.tabPage3.SuspendLayout();
            this.tabPage4.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnTesKoneksi
            // 
            this.btnTesKoneksi.Location = new System.Drawing.Point(12, 12);
            this.btnTesKoneksi.Name = "btnTesKoneksi";
            this.btnTesKoneksi.Size = new System.Drawing.Size(75, 23);
            this.btnTesKoneksi.TabIndex = 0;
            this.btnTesKoneksi.Text = "Tes Koneksi";
            this.btnTesKoneksi.UseVisualStyleBackColor = true;
            this.btnTesKoneksi.Click += new System.EventHandler(this.btnTesKoneksi_Click);
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Controls.Add(this.tabPage4);
            this.tabControl1.Location = new System.Drawing.Point(12, 41);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(411, 260);
            this.tabControl1.TabIndex = 1;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.lvwMahasiswa);
            this.tabPage1.Controls.Add(this.btnTampilkanData);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(403, 234);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "SELECT";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // lvwMahasiswa
            // 
            this.lvwMahasiswa.HideSelection = false;
            this.lvwMahasiswa.Location = new System.Drawing.Point(6, 35);
            this.lvwMahasiswa.Name = "lvwMahasiswa";
            this.lvwMahasiswa.Size = new System.Drawing.Size(391, 193);
            this.lvwMahasiswa.TabIndex = 1;
            this.lvwMahasiswa.UseCompatibleStateImageBehavior = false;
            // 
            // btnTampilkanData
            // 
            this.btnTampilkanData.Location = new System.Drawing.Point(6, 6);
            this.btnTampilkanData.Name = "btnTampilkanData";
            this.btnTampilkanData.Size = new System.Drawing.Size(112, 23);
            this.btnTampilkanData.TabIndex = 0;
            this.btnTampilkanData.Text = "Tampilkan Data";
            this.btnTampilkanData.UseVisualStyleBackColor = true;
            this.btnTampilkanData.Click += new System.EventHandler(this.btnTampilkanData_Click);
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.btnInsert);
            this.tabPage2.Controls.Add(this.txtAngkatanInsert);
            this.tabPage2.Controls.Add(this.txtNamaInsert);
            this.tabPage2.Controls.Add(this.label3);
            this.tabPage2.Controls.Add(this.txtNpmInsert);
            this.tabPage2.Controls.Add(this.label2);
            this.tabPage2.Controls.Add(this.label1);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(403, 234);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "INSERT";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // btnInsert
            // 
            this.btnInsert.Location = new System.Drawing.Point(65, 86);
            this.btnInsert.Name = "btnInsert";
            this.btnInsert.Size = new System.Drawing.Size(75, 23);
            this.btnInsert.TabIndex = 3;
            this.btnInsert.Text = "Insert";
            this.btnInsert.UseVisualStyleBackColor = true;
            this.btnInsert.Click += new System.EventHandler(this.btnInsert_Click);
            // 
            // txtAngkatanInsert
            // 
            this.txtAngkatanInsert.Location = new System.Drawing.Point(65, 60);
            this.txtAngkatanInsert.Name = "txtAngkatanInsert";
            this.txtAngkatanInsert.Size = new System.Drawing.Size(75, 20);
            this.txtAngkatanInsert.TabIndex = 2;
            // 
            // txtNamaInsert
            // 
            this.txtNamaInsert.Location = new System.Drawing.Point(65, 34);
            this.txtNamaInsert.Name = "txtNamaInsert";
            this.txtNamaInsert.Size = new System.Drawing.Size(332, 20);
            this.txtNamaInsert.TabIndex = 1;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(6, 63);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 13);
            this.label3.TabIndex = 0;
            this.label3.Text = "Angkatan";
            // 
            // txtNpmInsert
            // 
            this.txtNpmInsert.Location = new System.Drawing.Point(65, 6);
            this.txtNpmInsert.Name = "txtNpmInsert";
            this.txtNpmInsert.Size = new System.Drawing.Size(75, 20);
            this.txtNpmInsert.TabIndex = 0;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 37);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(35, 13);
            this.label2.TabIndex = 0;
            this.label2.Text = "Nama";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(31, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "NPM";
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.btnCariUpdate);
            this.tabPage3.Controls.Add(this.btnUpdate);
            this.tabPage3.Controls.Add(this.txtAngkatanUpdate);
            this.tabPage3.Controls.Add(this.label6);
            this.tabPage3.Controls.Add(this.txtNamaUpdate);
            this.tabPage3.Controls.Add(this.label5);
            this.tabPage3.Controls.Add(this.txtNpmUpdate);
            this.tabPage3.Controls.Add(this.label4);
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Size = new System.Drawing.Size(403, 234);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "UPDATE";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // btnCariUpdate
            // 
            this.btnCariUpdate.Location = new System.Drawing.Point(144, 4);
            this.btnCariUpdate.Name = "btnCariUpdate";
            this.btnCariUpdate.Size = new System.Drawing.Size(75, 23);
            this.btnCariUpdate.TabIndex = 1;
            this.btnCariUpdate.Text = "Cari";
            this.btnCariUpdate.UseVisualStyleBackColor = true;
            this.btnCariUpdate.Click += new System.EventHandler(this.btnCariUpdate_Click);
            // 
            // btnUpdate
            // 
            this.btnUpdate.Location = new System.Drawing.Point(63, 84);
            this.btnUpdate.Name = "btnUpdate";
            this.btnUpdate.Size = new System.Drawing.Size(75, 23);
            this.btnUpdate.TabIndex = 4;
            this.btnUpdate.Text = "Update";
            this.btnUpdate.UseVisualStyleBackColor = true;
            this.btnUpdate.Click += new System.EventHandler(this.btnUpdate_Click);
            // 
            // txtAngkatanUpdate
            // 
            this.txtAngkatanUpdate.Location = new System.Drawing.Point(63, 58);
            this.txtAngkatanUpdate.Name = "txtAngkatanUpdate";
            this.txtAngkatanUpdate.Size = new System.Drawing.Size(75, 20);
            this.txtAngkatanUpdate.TabIndex = 3;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(8, 61);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(53, 13);
            this.label6.TabIndex = 0;
            this.label6.Text = "Angkatan";
            // 
            // txtNamaUpdate
            // 
            this.txtNamaUpdate.Location = new System.Drawing.Point(63, 32);
            this.txtNamaUpdate.Name = "txtNamaUpdate";
            this.txtNamaUpdate.Size = new System.Drawing.Size(337, 20);
            this.txtNamaUpdate.TabIndex = 2;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(8, 35);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(35, 13);
            this.label5.TabIndex = 0;
            this.label5.Text = "Nama";
            // 
            // txtNpmUpdate
            // 
            this.txtNpmUpdate.Location = new System.Drawing.Point(63, 6);
            this.txtNpmUpdate.Name = "txtNpmUpdate";
            this.txtNpmUpdate.Size = new System.Drawing.Size(75, 20);
            this.txtNpmUpdate.TabIndex = 0;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(8, 9);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(31, 13);
            this.label4.TabIndex = 0;
            this.label4.Text = "NPM";
            // 
            // tabPage4
            // 
            this.tabPage4.Controls.Add(this.btnCariDelete);
            this.tabPage4.Controls.Add(this.btnDelete);
            this.tabPage4.Controls.Add(this.label9);
            this.tabPage4.Controls.Add(this.label8);
            this.tabPage4.Controls.Add(this.label7);
            this.tabPage4.Controls.Add(this.txtAngkatanDelete);
            this.tabPage4.Controls.Add(this.txtNamaDelete);
            this.tabPage4.Controls.Add(this.txtNpmDelete);
            this.tabPage4.Location = new System.Drawing.Point(4, 22);
            this.tabPage4.Name = "tabPage4";
            this.tabPage4.Size = new System.Drawing.Size(403, 234);
            this.tabPage4.TabIndex = 3;
            this.tabPage4.Text = "DELETE";
            this.tabPage4.UseVisualStyleBackColor = true;
            // 
            // btnCariDelete
            // 
            this.btnCariDelete.Location = new System.Drawing.Point(143, 5);
            this.btnCariDelete.Name = "btnCariDelete";
            this.btnCariDelete.Size = new System.Drawing.Size(75, 23);
            this.btnCariDelete.TabIndex = 1;
            this.btnCariDelete.Text = "Cari";
            this.btnCariDelete.UseVisualStyleBackColor = true;
            this.btnCariDelete.Click += new System.EventHandler(this.btnCariDelete_Click);
            // 
            // btnDelete
            // 
            this.btnDelete.Location = new System.Drawing.Point(62, 85);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(75, 23);
            this.btnDelete.TabIndex = 4;
            this.btnDelete.Text = "Delete";
            this.btnDelete.UseVisualStyleBackColor = true;
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(3, 62);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(53, 13);
            this.label9.TabIndex = 1;
            this.label9.Text = "Angkatan";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(3, 36);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(35, 13);
            this.label8.TabIndex = 1;
            this.label8.Text = "Nama";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(7, 10);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(31, 13);
            this.label7.TabIndex = 1;
            this.label7.Text = "NPM";
            // 
            // txtAngkatanDelete
            // 
            this.txtAngkatanDelete.Enabled = false;
            this.txtAngkatanDelete.Location = new System.Drawing.Point(62, 59);
            this.txtAngkatanDelete.Name = "txtAngkatanDelete";
            this.txtAngkatanDelete.Size = new System.Drawing.Size(75, 20);
            this.txtAngkatanDelete.TabIndex = 3;
            // 
            // txtNamaDelete
            // 
            this.txtNamaDelete.Enabled = false;
            this.txtNamaDelete.Location = new System.Drawing.Point(62, 33);
            this.txtNamaDelete.Name = "txtNamaDelete";
            this.txtNamaDelete.Size = new System.Drawing.Size(338, 20);
            this.txtNamaDelete.TabIndex = 2;
            // 
            // txtNpmDelete
            // 
            this.txtNpmDelete.Location = new System.Drawing.Point(62, 7);
            this.txtNpmDelete.Name = "txtNpmDelete";
            this.txtNpmDelete.Size = new System.Drawing.Size(75, 20);
            this.txtNpmDelete.TabIndex = 0;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(430, 306);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.btnTesKoneksi);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Latihan ADO.NET";
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.tabPage3.ResumeLayout(false);
            this.tabPage3.PerformLayout();
            this.tabPage4.ResumeLayout(false);
            this.tabPage4.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnTesKoneksi;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.TabPage tabPage4;
        private System.Windows.Forms.ListView lvwMahasiswa;
        private System.Windows.Forms.Button btnTampilkanData;
        private System.Windows.Forms.Button btnInsert;
        private System.Windows.Forms.TextBox txtAngkatanInsert;
        private System.Windows.Forms.TextBox txtNamaInsert;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtNpmInsert;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtNpmUpdate;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtAngkatanUpdate;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtNamaUpdate;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button btnCariUpdate;
        private System.Windows.Forms.Button btnUpdate;
        private System.Windows.Forms.Button btnCariDelete;
        private System.Windows.Forms.Button btnDelete;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtAngkatanDelete;
        private System.Windows.Forms.TextBox txtNamaDelete;
        private System.Windows.Forms.TextBox txtNpmDelete;
    }
}

