using System;
using System.Data;
using System.Windows.Forms;

using System.Data.OleDb;

namespace LatihanADONET
{
    public partial class Form1 : Form
    {
        // constructor
        public Form1()
        {
            InitializeComponent();
            InisialisasiListView();
        }

        private void btnTesKoneksi_Click(object sender, EventArgs e)
        {
            OleDbConnection conn = GetOpenConnection();

            if (conn.State == ConnectionState.Open)
            {
                MessageBox.Show("Koneksi ke database berhasil !",
                    "Informasi",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            else MessageBox.Show("Koneksi ke database gagal !!!",
                "Informasi",
                MessageBoxButtons.OK,
                MessageBoxIcon.Exclamation);
            conn.Dispose();
        }

        private OleDbConnection GetOpenConnection() {
            OleDbConnection conn = null;
            try
            {
                string dbName = @"D:\C\Matkul\Semester 3\Pemrograman lanjut\6\DbPerpustakaan.mdb";
                string connectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}", dbName);
                conn = new OleDbConnection(connectionString);
                conn.Open();
            }
            catch (Exception ex){
                MessageBox.Show("Error: " + ex.Message, 
                    "Error", 
                    MessageBoxButtons.OK, 
                    MessageBoxIcon.Error);
            }
            return conn;
        }

        private void InisialisasiListView()
        {
            lvwMahasiswa.View = View.Details;
            lvwMahasiswa.FullRowSelect = true;
            lvwMahasiswa.GridLines = true;

            lvwMahasiswa.Columns.Add("No.", 30, HorizontalAlignment.Center);
            lvwMahasiswa.Columns.Add("NPM", 70, HorizontalAlignment.Center);
            lvwMahasiswa.Columns.Add("Nama", 190, HorizontalAlignment.Left);
            lvwMahasiswa.Columns.Add("Angkatan", 70, HorizontalAlignment.Center);
        }

        private void btnTampilkanData_Click(object sender, EventArgs e)
        {
            lvwMahasiswa.Items.Clear();
            OleDbConnection conn = GetOpenConnection();
            string sql = @"select npm, nama, angkatan from mahasiswa order by nama";
            OleDbCommand cmd = new OleDbCommand(sql, conn);
            OleDbDataReader dtr = cmd.ExecuteReader();

            while (dtr.Read())
            {
                var noUrut = lvwMahasiswa.Items.Count + 1;

                var item = new ListViewItem(noUrut.ToString());
                item.SubItems.Add(dtr["npm"].ToString());
                item.SubItems.Add(dtr["nama"].ToString());
                item.SubItems.Add(dtr["angkatan"].ToString());

                lvwMahasiswa.Items.Add(item);
            }

            dtr.Dispose();
            cmd.Dispose();
            conn.Dispose();
        }

        private void btnInsert_Click(object sender, EventArgs e)
        {
            var result = 0;

            if(txtNpmInsert.Text.Length == 0)
            {
                MessageBox.Show("NPM harus diisi !!!", "Informasi", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtNamaInsert.Focus();
                return;
            }

            OleDbConnection conn = GetOpenConnection();
            var sql = @"insert into mahasiswa (npm, nama, angkatan)
                        values (@npm, @nama, @angkatan)";
            OleDbCommand cmd = new OleDbCommand(sql, conn);
            try
            {
                cmd.Parameters.AddWithValue("@npm", txtNpmInsert.Text);
                cmd.Parameters.AddWithValue("@nama", txtNamaInsert.Text);
                cmd.Parameters.AddWithValue("@angkatan", txtAngkatanInsert.Text);

                result = cmd.ExecuteNonQuery();
            }
            catch(Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                cmd.Dispose(); 
            }
            if (result > 0)
            {
                MessageBox.Show("Data mahasiswa berhasil disimpan !",
                    "Information",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                txtNpmInsert.Clear();
                txtNamaInsert.Clear();
                txtAngkatanInsert.Clear();
                txtNpmInsert.Focus();
            }
            else MessageBox.Show("Data mahasiswa gagal disimpan !!!", 
                "Informasi", 
                MessageBoxButtons.OK, 
                MessageBoxIcon.Exclamation);

            conn.Dispose();
        }

        private void btnCariUpdate_Click(object sender, EventArgs e)
        {
            if(txtNpmUpdate.Text.Length == 0)
            {
                MessageBox.Show("NPM harus !!!",
                    "Informasi",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Exclamation);
                txtNpmUpdate.Focus();
                return;
            }
            OleDbConnection conn = GetOpenConnection();
            string sql = @"select npm, nama, angkatan
                            from mahasiswa
                            where npm = @npm";

            OleDbCommand cmd = new OleDbCommand(sql, conn);
            cmd.Parameters.AddWithValue("@npm", txtNpmUpdate.Text);

            OleDbDataReader dtr = cmd.ExecuteReader();

            if (dtr.Read())
            {
                txtNpmUpdate.Text = dtr["npm"].ToString();
                txtNamaUpdate.Text = dtr["nama"].ToString();
                txtAngkatanUpdate.Text = dtr["angkatan"].ToString();
            }
            else MessageBox.Show("Data mahasiswa tidak ditemukan !",
                "Informasi",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            dtr.Dispose();
            cmd.Dispose();
            conn.Dispose();
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            var result = 0;

            if(txtNpmUpdate.Text.Length == 0)
            {
                MessageBox.Show("NPM harus !!!",
                    "Informasi",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Exclamation);
                txtNpmUpdate.Focus();
                return;
            }

            if(txtNamaUpdate.Text.Length == 0)
            {
                MessageBox.Show("Nama harus !!!", 
                    "Informasi", 
                    MessageBoxButtons.OK, 
                    MessageBoxIcon.Exclamation);
                txtNamaUpdate.Focus();
                return;
            }

            OleDbConnection conn = GetOpenConnection();

            string sql = @"update mahasiswa set nama = @nama, angkatan = @angkatan where npm = @npm";

            OleDbCommand cmd = new OleDbCommand(sql, conn);

            try
            {
                cmd.Parameters.AddWithValue("@nama", txtNamaUpdate.Text);
                cmd.Parameters.AddWithValue("@angkatan", txtAngkatanUpdate.Text);
                cmd.Parameters.AddWithValue("@npm", txtNpmUpdate.Text);

                result = cmd.ExecuteNonQuery();
            }
            catch(Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message,
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                cmd.Dispose();
            }
            if (result > 0)
            {
                MessageBox.Show("Data mahasiswa berhasil diupdate !",
                    "Informasi",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Exclamation);

                txtNpmUpdate.Clear();
                txtNamaUpdate.Clear();
                txtAngkatanUpdate.Clear();
                txtNpmUpdate.Focus();
            }
            else MessageBox.Show("Data mahasiswa gagal diupdate !!!", "Informasi", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            conn.Dispose();
        }

        private void btnCariDelete_Click(object sender, EventArgs e)
        {
            if (txtNpmDelete.Text.Length == 0)
            {
                MessageBox.Show("NPM harus !!!",
                    "Informasi",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Exclamation);
                txtNpmDelete.Focus();
                return;
            }
            OleDbConnection conn = GetOpenConnection();
            string sql = @"select npm, nama, angkatan
                            from mahasiswa
                            where npm = @npm";

            OleDbCommand cmd = new OleDbCommand(sql, conn);
            cmd.Parameters.AddWithValue("@npm", txtNpmDelete.Text);

            OleDbDataReader dtr = cmd.ExecuteReader();

            if (dtr.Read())
            {
                txtNpmDelete.Text = dtr["npm"].ToString();
                txtNamaDelete.Text = dtr["nama"].ToString();
                txtAngkatanDelete.Text = dtr["angkatan"].ToString();
            }
            else MessageBox.Show("Data mahasiswa tidak ditemukan !",
                "Informasi",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            dtr.Dispose();
            cmd.Dispose();
            conn.Dispose();
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            var result = 0;


            if (txtNpmDelete.Text.Length == 0)
            {
                MessageBox.Show("NPM harus !!!",
                    "Informasi",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Exclamation);
                txtNpmDelete.Focus();
                return;
            }

            if (txtNamaDelete.Text.Length == 0)
            {
                MessageBox.Show("Nama harus !!!",
                    "Informasi",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Exclamation);
                txtNamaDelete.Focus();
                return;
            }

            OleDbConnection conn = GetOpenConnection();

            var konfirmasi = MessageBox.Show("Apakah data mahasiswa ingin dihapus?", "Konfirmasi", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);

            if (konfirmasi == DialogResult.Yes)
            { 
                string sql = @"delete from mahasiswa where npm = @npm";

                OleDbCommand cmd = new OleDbCommand(sql, conn);

                try
                {
                    cmd.Parameters.AddWithValue("@npm", txtNpmDelete.Text);

                    result = cmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message,
                        "Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
                finally
                {
                    cmd.Dispose();
                }

                if (result > 0)
                {
                    MessageBox.Show("Data mahasiswa berhasil dihapus !",
                        "Informasi",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Exclamation);

                    txtNpmUpdate.Clear();
                    txtNamaUpdate.Clear();
                    txtAngkatanUpdate.Clear();
                    txtNpmUpdate.Focus();
                }
                else MessageBox.Show("Data mahasiswa gagal dihapus !!!", "Informasi", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            conn.Dispose();
        }
    }
}
