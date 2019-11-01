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
                MessageBox.Show("Koneksi Berhasil", "Informasi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else {
                MessageBox.Show("Koneksi Gagal", "Informasi", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            conn.Dispose();
        }

        private void btnTampilkanData_Click(object sender, EventArgs e)
        {
            lvwMahasiswa.Items.Clear();
            OleDbConnection conn = GetOpenConnection();
            string query = @"select npm, nama, angkatan 
                    from mahasiswa 
                    order by nama";
            OleDbCommand cmd = new OleDbCommand(query, conn);
            OleDbDataReader oddr = cmd.ExecuteReader();

            while (oddr.Read()) {
                var noUrut = lvwMahasiswa.Items.Count + 1;

                var item = new ListViewItem(noUrut.ToString());
                item.SubItems.Add(oddr["npm"].ToString());
                item.SubItems.Add(oddr["nama"].ToString());
                item.SubItems.Add(oddr["angkatan"].ToString());

                lvwMahasiswa.Items.Add(item);
            }

            oddr.Dispose();
            cmd.Dispose();
            conn.Dispose();
        }

        private void btnInsert_Click(object sender, EventArgs e)
        {
            var result = 0;

            // validasi npm harus diisi
            if (txtNpmInsert.Text.Length == 0)
            {
                MessageBox.Show("NPM harus diisi !!!", "Informasi", MessageBoxButtons.OK,
                        MessageBoxIcon.Exclamation);

                txtNpmInsert.Focus();
                return;
            }

            // validasi nama harus diisi
            if (txtNamaInsert.Text.Length == 0)
            {
                MessageBox.Show("Nama harus diisi !!!", "Informasi", MessageBoxButtons.OK,
                        MessageBoxIcon.Exclamation);

                txtNamaInsert.Focus();
                return;
            }

            // membuat objek Connection, sekaligus buka koneksi ke database
            OleDbConnection conn = GetOpenConnection();

            // deklarasi variabel sql untuk menampung perintah INSERT
            var sql = @"insert into mahasiswa (npm, nama, angkatan)
                values (@npm, @nama, @angkatan)";

            // membuat objek Command untuk mengeksekusi perintah SQL
            OleDbCommand cmd = new OleDbCommand(sql, conn);
            try
            {
                // set parameter untuk nama, angkatan dan npm
                cmd.Parameters.AddWithValue("@npm", txtNpmInsert.Text);
                cmd.Parameters.AddWithValue("@nama", txtNamaInsert.Text);
                cmd.Parameters.AddWithValue("@angkatan", txtAngkatanInsert.Text);

                result = cmd.ExecuteNonQuery(); // eksekusi perintah INSERT
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
            }
            finally
            {
                cmd.Dispose();
            }

            if (result > 0)
            {
                MessageBox.Show("Data mahasiswa berhasil disimpan !", "Informasi", MessageBoxButtons.OK,
                        MessageBoxIcon.Information);

                // reset form
                txtNpmInsert.Clear();
                txtNamaInsert.Clear();
                txtAngkatanInsert.Clear();
                txtNpmInsert.Focus();
            }
            else
                MessageBox.Show("Data mahasiswa gagal disimpan !!!", "Informasi", MessageBoxButtons.OK,
                        MessageBoxIcon.Exclamation);

            // setelah selesai digunakan, 
            // segera hapus objek connection dari memory
            conn.Dispose();
        }

        private void btnCariUpdate_Click(object sender, EventArgs e)
        {
            if (txtNpmUpdate.Text.Length == 0){
                MessageBox.Show("NPM harus !!!", "Informasi", MessageBoxButtons.OK,
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
           
            OleDbDataReader dtr = cmd.ExecuteReader(); // eksekusi perintah SELECT

            if (dtr.Read()) // data ditemukan
            {
                txtNpmUpdate.Text = dtr["npm"].ToString();
                txtNamaUpdate.Text = dtr["nama"].ToString();
                txtAngkatanUpdate.Text = dtr["angkatan"].ToString();
            }
            else {
                 MessageBox.Show("Data mahasiswa tidak ditemukan !", "Informasi", MessageBoxButtons.OK,
                                        MessageBoxIcon.Information);
            }
           
            dtr.Dispose();
            cmd.Dispose();
            conn.Dispose();
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            var result = 0;

            if (txtNpmUpdate.Text.Length == 0)
            {
                MessageBox.Show("NPM harus !!!", "Informasi", MessageBoxButtons.OK,
                            MessageBoxIcon.Exclamation);

                txtNpmUpdate.Focus();
                return;
            }
         
            if (txtNamaUpdate.Text.Length == 0)
            {
                MessageBox.Show("Nama harus !!!", "Informasi", MessageBoxButtons.OK,
                            MessageBoxIcon.Exclamation);

                txtNamaUpdate.Focus();
                return;
            }
            
            OleDbConnection conn = GetOpenConnection();
            string sql = @"update mahasiswa set nama = @nama, angkatan = @angkatan
                    where npm = @npm";
            OleDbCommand cmd = new OleDbCommand(sql, conn);

            try
            {
                // set parameter untuk nama, angkatan dan npm
                cmd.Parameters.AddWithValue("@nama", txtNamaUpdate.Text);
                cmd.Parameters.AddWithValue("@angkatan", txtAngkatanUpdate.Text);
                cmd.Parameters.AddWithValue("@npm", txtNpmUpdate.Text);

                result = cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
            }
            finally
            {
                cmd.Dispose();
            }

            if (result > 0)
            {
                MessageBox.Show("Data mahasiswa berhasil diupdate !", "Informasi", MessageBoxButtons.OK,
                        MessageBoxIcon.Information);

                // reset form
                txtNpmUpdate.Clear();
                txtNamaUpdate.Clear();
                txtAngkatanUpdate.Clear();
                txtNpmUpdate.Focus();
            }
            else
                MessageBox.Show("Data mahasiswa gagal diupdate !!!", "Informasi", MessageBoxButtons.OK,
                        MessageBoxIcon.Exclamation);
            
            conn.Dispose();
        }

        private void btnCariDelete_Click(object sender, EventArgs e)
        {
            if (txtNpmDelete.Text.Length == 0)
            {
                MessageBox.Show("NPM harus !!!", "Informasi", MessageBoxButtons.OK,
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

            OleDbDataReader dtr = cmd.ExecuteReader(); // eksekusi perintah SELECT

            if (dtr.Read()) // data ditemukan
            {
                txtNpmDelete.Text = dtr["npm"].ToString();
                txtNamaDelete.Text = dtr["nama"].ToString();
                txtAngkatanDelete.Text = dtr["angkatan"].ToString();
            }
            else
            {
                MessageBox.Show("Data mahasiswa tidak ditemukan !", "Informasi", MessageBoxButtons.OK,
                                       MessageBoxIcon.Information);
            }

            dtr.Dispose();
            cmd.Dispose();
            conn.Dispose();
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            var result = 0;

            if (txtNpmDelete.Text.Length == 0)
            {
                MessageBox.Show("NPM harus !!!", "Informasi", MessageBoxButtons.OK,
                            MessageBoxIcon.Exclamation);

                txtNpmDelete.Focus();
                return;
            }

            if (txtNamaDelete.Text.Length == 0)
            {
                MessageBox.Show("Nama harus !!!", "Informasi", MessageBoxButtons.OK,
                            MessageBoxIcon.Exclamation);

                txtNpmDelete.Focus();
                return;
            }

            OleDbConnection conn = GetOpenConnection();
            string sql = @"delete from mahasiswa where npm = @npm";
            OleDbCommand cmd = new OleDbCommand(sql, conn);
           

            try
            {
                cmd.Parameters.AddWithValue("@npm", txtNpmDelete.Text);
                result = cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
            }
            finally
            {
                cmd.Dispose();
            }

            if (result > 0)
            {
                MessageBox.Show("Data mahasiswa berhasil dihapus !", "Informasi", MessageBoxButtons.OK,
                        MessageBoxIcon.Information);

                // reset form
                txtNpmDelete.Clear();
                txtNamaDelete.Clear();
                txtAngkatanDelete.Clear();
                txtNpmDelete.Focus();
            }
            else
                MessageBox.Show("Data mahasiswa gagal dihapus !!!", "Informasi", MessageBoxButtons.OK,
                        MessageBoxIcon.Exclamation);

            conn.Dispose();
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


        private OleDbConnection GetOpenConnection() {
            OleDbConnection conn = null;

            try
            {
                string dbName = @"D:\18.11.2288\LatihanADO.NET\Database\DbPerpustakaan.mdb";
                string connString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + dbName);
                conn = new OleDbConnection(connString);
                conn.Open();
            }
            catch (Exception e){
                MessageBox.Show("Error: " + e.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return conn;
        }
    }
}
