using System;
using System.Data.OleDb;
using System.IO;
using System.Windows.Forms;

namespace WindowsFormsApp3
{
    public partial class Form4 : Form
    {
        private string exeDir;
        private string dbPath;
        private string myconn;
        private OleDbConnection conn;

        public Form4()
        {
            InitializeComponent();

            exeDir = AppDomain.CurrentDomain.BaseDirectory;
            dbPath = Path.Combine(exeDir, "Database11.accdb");

            myconn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dbPath + ";Persist Security Info=False;";
            conn = new OleDbConnection(myconn);
        }

        private void load()
        {
            string sel = "SELECT * FROM Zap";
            try
            {
                conn.Open();
                OleDbCommand select = new OleDbCommand(sel, conn);
                OleDbDataReader reader = select.ExecuteReader();

                dataGridView1.Rows.Clear();
                dataGridView1.Columns.Clear();
                for (int i = 0; i < reader.FieldCount; i++)
                    dataGridView1.Columns.Add(reader.GetName(i), reader.GetName(i));

                while (reader.Read())
                {
                    object[] row = new object[reader.FieldCount];
                    reader.GetValues(row);
                    dataGridView1.Rows.Add(row);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                conn.Close();
            }
        }
        private void Addrow()
        {
            if (dataGridView1.SelectedRows.Count == 0)
                return;
            try
            {
                conn.Open();

                foreach (DataGridViewRow row in dataGridView1.SelectedRows)
                {
                    string insertQuery = "INSERT INTO Zap ([Имя], [Телефон], [Дата], [Время]) VALUES (?, ?, ?, ?)";
                    using (OleDbCommand cmd = new OleDbCommand(insertQuery, conn))
                    {
                        cmd.Parameters.AddWithValue("?", row.Cells["Имя"].Value ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("?", row.Cells["Телефон"].Value ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("?", row.Cells["Дата"].Value ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("?", row.Cells["Время"].Value ?? DBNull.Value);

                        cmd.ExecuteNonQuery();
                    }
                }

                MessageBox.Show("Данные добавлены!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при добавлении: " + ex.Message);
            }
            finally
            {
                conn.Close();
            }
            load();
        }
        private void Delrow()
        {
            if (dataGridView1.SelectedRows.Count == 0)
                return;
            try
            {
                conn.Open();
                foreach (DataGridViewRow row in dataGridView1.SelectedRows)
                {
                    string deleteQuery = "DELETE FROM Zap WHERE [Код] = ?";

                    using (OleDbCommand cmd = new OleDbCommand(deleteQuery, conn))
                    {
                        cmd.Parameters.AddWithValue("?", row.Cells["Код"].Value);
                        cmd.ExecuteNonQuery();
                    }
                }
                MessageBox.Show("Данные удалены!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при удалении: " + ex.Message);
            }
            finally
            {
                conn.Close();
            }
            load();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Addrow();
            label7.Text = "Вы успешно записаны!";
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
          
        }

        private void Form4_Load(object sender, EventArgs e)
        {
            load();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form3 f = new Form3();
            f.Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Delrow();
        }
    }
}
