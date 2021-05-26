using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KursovayaKP
{
    public partial class Sostav : Form
    {
        public Form1 f1;
        OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=E:\Практика\Tecn.accdb");
        public Sostav()
        {
            InitializeComponent();
        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void Sostav_Load(object sender, EventArgs e)
        {
            if (listBox1.Items.Count == 0)
            {
                string[] array = getMaterial().Select(n => n.ToString()).ToArray();
                listBox1.Items.AddRange(array);

            }

        }
        private int getIdByMaterial(string nameTovar)
        {
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT Код_Материала FROM Материалы where Наименование = '" + nameTovar + "'";
            cmd.ExecuteNonQuery();
            con.Close();
            DataTable dt = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);
            return int.Parse(dt.Rows[0][0].ToString());
        }
        private List<string> getMaterial()
        {
            List<string> opers = new List<string>();
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT * FROM Материалы";
            cmd.ExecuteNonQuery();
            con.Close();
            DataTable dt = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);
            foreach (DataRow item in dt.Rows)
            {
                opers.Add(item[1].ToString());
            }
            return opers;
        }

        private void bunifuImageButton3_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedItem != null)
            {
                listBox2.Items.Add(listBox1.SelectedItem);
                listBox1.Items.Remove(listBox1.SelectedItem);
            }
            else { MessageBox.Show("Не выбраны элементы", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Information); }
        }

        private void bunifuImageButton5_Click(object sender, EventArgs e)
        {
            Form2 f2 = new Form2(f1);
            string[] aa = bunifuCustomLabel1.Text.Split(' ');
            int nn = int.Parse(aa[1]);
            string[] kolvo = richTextBox1.Text.Split('\n');
            for (int i = 0; i <= listBox2.Items.Count - 1; i++)
            {
                int IdTov = getIdByMaterial(listBox2.Items[i].ToString());
                int B = int.Parse(kolvo[i]);
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "INSERT INTO Состав_заказа (Код_Материала, [Кол-во], Id_Заказа) VALUES("+IdTov+","+B+","+ nn +")";
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT Состав_Заказа.Id_Состава_Заказа, Материалы.Наименование, Состав_Заказа.[Кол-во] FROM Материалы INNER JOIN Состав_Заказа ON Материалы.Код_Материала = Состав_Заказа.Код_Материала WHERE Id_Заказа = " + nn + "";
                cmd.ExecuteNonQuery();
                con.Close();
                DataTable dt = new DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                bunifuDataGridView1.DataSource = dt;
            }

            for (int i = 0; i < listBox2.Items.Count; i++)
            {
                listBox1.Items.Add(listBox2.Items[i]);
            }
            listBox2.Items.Clear();
            richTextBox1.Clear();
        }

        private void bunifuImageButton1_Click(object sender, EventArgs e)
        {
            this.Hide();
        }

        private void bunifuImageButton6_Click(object sender, EventArgs e)
        {
            string[] aa = bunifuCustomLabel1.Text.Split(' ');
            int nn = int.Parse(aa[1]);
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = $"DELETE FROM Состав_заказа WHERE Id_Состава_Заказа = {bunifuDataGridView1[0, bunifuDataGridView1.CurrentCell.RowIndex].Value}";
            cmd.ExecuteNonQuery();
            cmd.CommandText = "SELECT Состав_Заказа.Id_Состава_Заказа, Материалы.Наименование, Состав_Заказа.[Кол-во] FROM Материалы INNER JOIN Состав_Заказа ON Материалы.Код_Материала = Состав_Заказа.Код_Материала WHERE Id_Заказа = " + nn + "";
            cmd.ExecuteNonQuery();
            con.Close();
            DataTable dt = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);
            bunifuDataGridView1.DataSource = dt;
        }

        private void bunifuImageButton4_Click(object sender, EventArgs e)
        {
            if (listBox2.SelectedItem != null)
            {
                listBox1.Items.Add(listBox2.SelectedItem);
                listBox2.Items.Remove(listBox2.SelectedItem);
            }

            else { MessageBox.Show("Не выбраны элементы", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Information); }
        }
    }
}
