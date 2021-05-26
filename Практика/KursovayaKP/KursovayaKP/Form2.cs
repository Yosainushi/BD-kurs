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
using Application = Microsoft.Office.Interop.Excel.Application;
using Word = Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.IO;
using Microsoft.Office.Interop.Word;

namespace KursovayaKP
{
    public partial class Form2 : Form
    {
        private Application xlExcel;
        private Workbook xlWorkBook;
        Word._Application oWord = new Word.Application();
        public Form1 f1;
        OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=E:\Практика\Tecn.accdb");
        public Form2(Form1 f1)
        {
            
            InitializeComponent();
            bunifuFlatButton1.Normalcolor = System.Drawing.Color.FromArgb(26, 32, 50);
            bunifuFlatButton2.Normalcolor = System.Drawing.Color.FromArgb(26, 32, 50);
            bunifuFlatButton3.Normalcolor = System.Drawing.Color.FromArgb(26, 32, 50);
            bunifuFlatButton4.Normalcolor = System.Drawing.Color.FromArgb(26, 32, 50);
            bunifuFlatButton1.BackColor = System.Drawing.Color.FromArgb(26, 32, 50);
            bunifuFlatButton2.BackColor = System.Drawing.Color.FromArgb(26, 32, 50);
            bunifuFlatButton3.BackColor = System.Drawing.Color.FromArgb(26, 32, 50);
            bunifuFlatButton4.BackColor = System.Drawing.Color.FromArgb(26, 32, 50);
            Login login = new Login();
            Form1 form1 = new Form1(login);
        
        }
        public Login login = new Login();
        public string text1;
        public string text2;
        
        
        public void loadTable(string selectTable)
        {
            Form1 f1 = new Form1(login);
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = selectTable;
            cmd.ExecuteNonQuery();
            con.Close();
            System.Data.DataTable dt = new System.Data.DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);
            bunifuDataGridView1.DataSource = dt;
        }
        public void Loading()
        {
            if (bunifuCustomLabel1.Text == "Заказы")
            {
                loadTable(Queries.selectZakazi);
                text1 = bunifuDropdown1.Text;
                text2 = bunifuDropdown2.Text;

            }
            if (bunifuCustomLabel1.Text == "Заказчики")
            {
                loadTable(Queries.selectZakazch);
                text1 = bunifuDropdown1.Text;
                text2 = bunifuDropdown2.Text;
            }
            if (bunifuCustomLabel1.Text == "Материалы")
            {
                loadTable(Queries.selectMat);
                text1 = bunifuDropdown1.Text;
                text2 = bunifuDropdown2.Text;
            }

            if (bunifuCustomLabel1.Text == "Поставки")
            {
                loadTable(Queries.selectPostavki);
                text1 = bunifuDropdown1.Text;
                text2 = bunifuDropdown2.Text;
            }
            if (bunifuCustomLabel1.Text == "Поставщики")
            {
                loadTable(Queries.selectPostavch);
                text1 = bunifuDropdown1.Text;
                text2 = bunifuDropdown2.Text;
            }
            if (bunifuCustomLabel1.Text == "Приказы")
            {
                loadTable(Queries.selectPrikazi);
                text1 = bunifuDropdown1.Text;
                text2 = bunifuDropdown2.Text;
            }
            if (bunifuCustomLabel1.Text == "Работники")
            {
                loadTable(Queries.selectRabotniki);
                text1 = bunifuDropdown1.Text;
                text2 = bunifuDropdown2.Text;
            }
            if (bunifuCustomLabel1.Text == "Склады")
            {
                loadTable(Queries.selectScladi);
                text1 = bunifuDropdown1.Text;
                text2 = bunifuDropdown2.Text;
            }
            else if (bunifuCustomLabel1.Text == "Пользователи")
            {

                loadTable(Queries.selectPolz);
                text1 = bunifuDropdown1.Text;
                text2 = bunifuDropdown2.Text;
            }
        }
        private void Form2_Load(object sender, EventArgs e)
        {
            foreach (DataGridViewColumn column in bunifuDataGridView1.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }

           
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT COUNT(*) FROM Пользователи WHERE Логин = '" + bunifuTextBox1.Text + "' AND Пароль = '" + bunifuTextBox2.Text + "' AND Право_Администрации= 'Да'";
            int x = int.Parse(cmd.ExecuteScalar().ToString());
            if (x != 0)
            {
                comboBox5.Items.Add("Пользователи");
            }
            con.Close();
            bunifuDropdown1.Items.Clear();
            bunifuDropdown2.Items.Clear();
            bunifuDropdown3.Items.Clear();
            bunifuDropdown4.Items.Clear();
            bunifuTextBox1.Visible = false;
            bunifuTextBox2.Visible = false;
            bunifuTextBox3.Visible = false;
            bunifuTextBox4.Visible = false;
            bunifuDropdown1.Visible = false;
            bunifuDropdown2.Visible = false;
            bunifuDropdown3.Visible = false;
            bunifuDropdown4.Visible = false;
            bunifuDatePicker2.Visible = false;
            checkBox1.Visible = false;
            comboBox5.DropDownStyle = ComboBoxStyle.DropDownList;
            if (bunifuCustomLabel1.Text == "Заказы")
            {
                loadTable(Queries.selectZakazi);
                bunifuDataGridView1.Columns[0].Visible = false;
                bunifuDataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                bunifuFlatButton5.Visible = true;
                bunifuTextBox1.Visible = true;
                bunifuTextBox2.Visible = true;
                bunifuDatePicker2.Visible = true;
                bunifuDropdown1.Visible = true;
                bunifuDropdown1.Text = "Заказчик";
                bunifuDropdown2.Text = "Оформляющий";
                bunifuDropdown2.Visible = true;
                bunifuFlatButton8.Visible = true;
                bunifuDropdown1.Items.AddRange(getClient().ToArray());
                bunifuDropdown2.Items.AddRange(getSotrudniki().ToArray());
            }
            if (bunifuCustomLabel1.Text == "Заказчики")
            {
                loadTable(Queries.selectZakazch);
                bunifuDataGridView1.Columns[0].Visible = false;
                bunifuDataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                bunifuTextBox1.Visible = true;
                bunifuTextBox2.Visible = true;
                bunifuTextBox3.Visible = true;
                bunifuTextBox4.Visible = true;
            }
            if (bunifuCustomLabel1.Text == "Материалы")
            {
                bunifuDropdown1.Visible = true;
                bunifuDropdown1.Text = "Склад";
                bunifuTextBox1.Visible = true;
                bunifuTextBox2.Visible = true;
                bunifuTextBox3.Visible = true;
                bunifuDropdown1.Items.AddRange(getSklad().ToArray());
                loadTable(Queries.selectMat);

                bunifuDataGridView1.Columns[0].Visible = false;
                bunifuDataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            }
            
            if (bunifuCustomLabel1.Text == "Поставки")
            {
                bunifuDropdown1.Text = "Поставщик";
                bunifuDropdown2.Text = "Материал";
                bunifuDropdown3.Text = "Адрес склада";
                bunifuDropdown4.Text = "Ответственный";
                bunifuDatePicker2.Visible = true;
                bunifuDropdown1.Visible = true;
                bunifuDropdown2.Visible = true;
                bunifuDropdown3.Visible = true;
                bunifuDropdown4.Visible = true;
                bunifuDropdown1.Items.AddRange(getPostavshikt().ToArray());
                bunifuDropdown2.Items.AddRange(getMaterial().ToArray());
                bunifuDropdown3.Items.AddRange(getSklad().ToArray());
                bunifuDropdown4.Items.AddRange(getSotrudnikiSkl().ToArray());
                loadTable(Queries.selectPostavki);
                bunifuDataGridView1.Columns[0].Visible = false;
                bunifuDataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            }
            if (bunifuCustomLabel1.Text == "Поставщики")
            {
                loadTable(Queries.selectPostavch);
                bunifuTextBox1.Visible = true;
                bunifuTextBox2.Visible = true;
                bunifuTextBox3.Visible = true;
                bunifuTextBox4.Visible = true;
                bunifuDataGridView1.Columns[0].Visible = false;
                bunifuDataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            }
           
            if (bunifuCustomLabel1.Text == "Работники")
            {
                bunifuDropdown1.Text = "Проффесия";
                bunifuTextBox1.Visible = true;
                bunifuTextBox2.Visible = true;
                bunifuTextBox3.Visible = true;
                bunifuTextBox4.Visible = true;
                bunifuDropdown1.Visible = true;
                string[] mas = { "Зав.Склада", "Менеджер"};
                bunifuDropdown1.Items.AddRange(mas);
                loadTable(Queries.selectRabotniki);
                bunifuDataGridView1.Columns[0].Visible = false;
                bunifuDataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            }
            if (bunifuCustomLabel1.Text == "Склады")
            {
                bunifuTextBox1.Visible = true;
                loadTable(Queries.selectScladi);
                bunifuDataGridView1.Columns[0].Visible = false;
                bunifuDataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            }
            else if (bunifuCustomLabel1.Text == "Пользователи")
            {

                loadTable(Queries.selectPolz);
                bunifuTextBox1.Visible = true;
                bunifuTextBox2.Visible = true;
                bunifuDropdown1.Text = "Право Админа";
                bunifuTextBox1.PlaceholderText = "Логин";
                bunifuTextBox2.PlaceholderText = "Пароль";
                bunifuDropdown1.Items.Add("Да");
                bunifuDropdown1.Items.Add("Нет");
                bunifuDropdown1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
                bunifuDropdown1.Visible = true;
                bunifuDataGridView1.Columns[0].Visible = false;
                bunifuDropdown1.Location = new System.Drawing.Point(667, 102);
                bunifuDataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            }
        }

        private void Form2_Shown(object sender, EventArgs e)
        {

        }


        private List<string> getPostavshikt()
        {
            List<string> opers = new List<string>();
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT * FROM Поставщики ";
            cmd.ExecuteNonQuery();
            con.Close();
            System.Data.DataTable dt = new System.Data.DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);
            foreach (DataRow item in dt.Rows)
            {
                opers.Add(item[1].ToString());
            }
            return opers;
        }
        private int getIdByPostavshik(string nameOper)
        {
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT Id_Поставщика FROM Поставщики where Название_организации = '" + nameOper + "'";
            cmd.ExecuteNonQuery();
            con.Close();
            System.Data.DataTable dt = new System.Data.DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);
            return int.Parse(dt.Rows[0][0].ToString());
        }
        private List<string> getClient()
        {
            List<string> opers = new List<string>();
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT * FROM Заказчики";
            cmd.ExecuteNonQuery();
            con.Close();
            System.Data.DataTable dt = new System.Data.DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);
            foreach (DataRow item in dt.Rows)
            {
                opers.Add(item[1].ToString() + " " + item[2].ToString() + " " + item[3].ToString());
            }
            return opers;
        }
        private int getIdByClient(string nameOper)
        {
            string[] a = nameOper.Split(' ');
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT Id_Заказчика FROM Заказчики where Фамилия = '" + a[0] + "' and Имя = '" + a[1] + "' and Отчество = '" + a[2] + "'";
            cmd.ExecuteNonQuery();
            con.Close();
            System.Data.DataTable dt = new System.Data.DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);
            return int.Parse(dt.Rows[0][0].ToString());
        }
        private List<string> getSklad()
        {
            List<string> opers = new List<string>();
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT * FROM Склады";
            cmd.ExecuteNonQuery();
            con.Close();
            System.Data.DataTable dt = new System.Data.DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);
            foreach (DataRow item in dt.Rows)
            {
                opers.Add(item[1].ToString());
            }
            return opers;
        }
        private int getIdBySklad(string nameOper)
        {
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT Id_Склада FROM Склады where Адрес_Склада = '" + nameOper + "' ";
            cmd.ExecuteNonQuery();
            con.Close();
            System.Data.DataTable dt = new System.Data.DataTable();
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
            System.Data.DataTable dt = new System.Data.DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);
            foreach (DataRow item in dt.Rows)
            {
                opers.Add(item[1].ToString());
            }
            return opers;
        }
        private int getIdByMaterial(string nameOper)
        {
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT Код_Материала FROM Материалы where Наименование = '" + nameOper + "'";
            cmd.ExecuteNonQuery();
            con.Close();
            System.Data.DataTable dt = new System.Data.DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);
            return int.Parse(dt.Rows[0][0].ToString());
        }
        private List<string> getSotrudniki()
        {
            List<string> opers = new List<string>();
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT * FROM Работники where Профессия='Менеджер'";
            cmd.ExecuteNonQuery();
            con.Close();
            System.Data.DataTable dt = new System.Data.DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);
            foreach (DataRow item in dt.Rows)
            {
                opers.Add(item[2].ToString()+' '+ item[3].ToString()+ ' ' +item[4].ToString());
            }
            return opers;
        }
        private List<string> getSotrudnikiSkl()
        {
            List<string> opers = new List<string>();
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT * FROM Работники where Профессия='Зав.Склада'";
            cmd.ExecuteNonQuery();
            con.Close();
            System.Data.DataTable dt = new System.Data.DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);
            foreach (DataRow item in dt.Rows)
            {
                opers.Add(item[2].ToString() + ' ' + item[3].ToString() + ' ' + item[4].ToString());
            }
            return opers;
        }
        private int getIdBySotrudniki(string nameOper)
        {
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT Id_Работника FROM Работники where Фамилия = '" + nameOper.Split(' ')[0]+ "' and Имя = '" + nameOper.Split(' ')[1] + "' and Отчество = '" + nameOper.Split(' ')[2] + "'";
            cmd.ExecuteNonQuery();
            con.Close();
            System.Data.DataTable dt = new System.Data.DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);
            return int.Parse(dt.Rows[0][0].ToString());
        }





        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            
        }

        

        private void button1_Click_1(object sender, EventArgs e)
        {

        }

        private void bunifuTextBox1_TextChanged(object sender, EventArgs e)
        {


        }

        private void button1_Click_2(object sender, EventArgs e)
        {
            this.Form2_Load(sender, e);
        }

        

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void bunifuTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (bunifuCustomLabel1.Text == "Заказчики")
            {
                if (!Char.IsLetter(e.KeyChar) && e.KeyChar != 8)
                {
                    e.Handled = true;
                }
            }
            else
            if (bunifuCustomLabel1.Text == "Налоги")
            {
                if (!Char.IsLetter(e.KeyChar) && e.KeyChar != 8 && e.KeyChar != ' ')
                {
                    e.Handled = true;
                }
            }
            else
            if (bunifuCustomLabel1.Text == "Налогоплательщики")
            {
                if (!Char.IsLetter(e.KeyChar) && e.KeyChar != 8 && e.KeyChar != '-')
                {
                    e.Handled = true;
                }
            }
            else
            if (bunifuCustomLabel1.Text == "Сотрудники")
            {
                if (!Char.IsLetter(e.KeyChar) && e.KeyChar != 8 && e.KeyChar != '-')
                {
                    e.Handled = true;
                }
            }
            else
            if (bunifuCustomLabel1.Text == "ВидОплаты")
            {
                if (!Char.IsLetter(e.KeyChar) && e.KeyChar != 8 && e.KeyChar != '-')
                {
                    e.Handled = true;
                }
            }
            if (bunifuCustomLabel1.Text == "Работники")
            {
                if (!Char.IsLetter(e.KeyChar) && e.KeyChar != 8)
                {
                    e.Handled = true;
                }
            }
        }

        private void bunifuTextBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (bunifuCustomLabel1.Text == "Заказчики")
            {
                if (!Char.IsLetter(e.KeyChar) && e.KeyChar != 8)
                {
                    e.Handled = true;
                }
            }
            if (bunifuCustomLabel1.Text == "Работники")
            {
                if (!Char.IsLetter(e.KeyChar) && e.KeyChar != 8)
                {
                    e.Handled = true;
                }
            }
            else
            if (bunifuCustomLabel1.Text == "Заказы")
            {
                if (!Char.IsDigit(e.KeyChar) && e.KeyChar != 8 && e.KeyChar != '-')
                {
                    e.Handled = true;
                }
            }
            else
            if (bunifuCustomLabel1.Text == "Материалы")
            {
                if (!Char.IsDigit(e.KeyChar) && e.KeyChar != 8 && e.KeyChar != '-')
                {
                    e.Handled = true;
                }
            }
            else
            if (bunifuCustomLabel1.Text == "Поставщики")
            {
                if (!Char.IsDigit(e.KeyChar) && e.KeyChar != 8 && e.KeyChar != '-')
                {
                    e.Handled = true;
                }
            }
        }

        private void bunifuTextBox3_KeyPress(object sender, KeyPressEventArgs e)
        {

            if (bunifuCustomLabel1.Text == "Заказчики")
            {
                if (!Char.IsLetter(e.KeyChar) && e.KeyChar != 8)
                {
                    e.Handled = true;
                }
            }
            if (bunifuCustomLabel1.Text == "Работники")
            {
                if (!Char.IsLetter(e.KeyChar) && e.KeyChar != 8)
                {
                    e.Handled = true;
                }
            }
            else
             if (bunifuCustomLabel1.Text == "Материалы")
            {
                if (!Char.IsDigit(e.KeyChar) && e.KeyChar != 8 && e.KeyChar != '-')
                {
                    e.Handled = true;
                }
            }
            if (bunifuCustomLabel1.Text == "Банки")
            {
                if (!Char.IsDigit(e.KeyChar) && e.KeyChar != 8 && e.KeyChar != '-')
                {
                    e.Handled = true;
                }
            }
            if (bunifuCustomLabel1.Text == "Почты")
            {
                if (!Char.IsDigit(e.KeyChar) && e.KeyChar != 8 && e.KeyChar != '-')
                {
                    e.Handled = true;
                }
            }
        }

        private void bunifuTextBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (bunifuCustomLabel1.Text == "Заказчики")
            {
                if (!Char.IsDigit(e.KeyChar) && e.KeyChar != 8)
                {
                    e.Handled = true;
                }
            }
            if (bunifuCustomLabel1.Text == "Работники")
            {
                if (!Char.IsDigit(e.KeyChar) && e.KeyChar != 8 && e.KeyChar != '-')
                {
                    e.Handled = true;
                }
            }
            if (bunifuCustomLabel1.Text == "Сотрудники")
            {
                if (!Char.IsDigit(e.KeyChar) && e.KeyChar != 8)
                {
                    e.Handled = true;
                }
            }
            if (bunifuCustomLabel1.Text == "Поставщики")
            {
                if (!Char.IsDigit(e.KeyChar) && e.KeyChar != 8 && e.KeyChar != '-')
                {
                    e.Handled = true;
                }
            }
        }

        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void bunifuTextBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void bunifuFlatButton3_Click(object sender, EventArgs e)
        {
           
                int ID = Convert.ToInt32(bunifuDataGridView1.Rows[bunifuDataGridView1.CurrentRow.Index].Cells[0].Value);
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Delete * FROM "+ bunifuCustomLabel1.Text +" WHERE "+ bunifuDataGridView1.Columns[0].HeaderText +"=" + ID + "";
                cmd.ExecuteNonQuery();
                con.Close();
                Loading();
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            



            bunifuTextBox1.Clear();
            bunifuTextBox2.Clear();
            bunifuTextBox3.Clear();
            bunifuTextBox4.Clear();
            bunifuDropdown1.Text = "";
            bunifuDropdown2.Text = "";
            bunifuDropdown3.Text = "";
            bunifuDropdown4.Text = "";
            comboBox5.Text = "";

        }

        private void bunifuImageButton1_Click(object sender, EventArgs e)
        {
            Form2_Load(sender, e);
        }

        private void bunifuFlatButton4_Click(object sender, EventArgs e)
        {
            if (comboBox5.Visible == true)
            {
                comboBox5.Visible = false;
            }
            else
            {
                comboBox5.Visible = true;
                comboBox5.Focus();
            }
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            bunifuTextBox1.Text = "";
            bunifuTextBox2.Text = "";
            bunifuTextBox3.Text = "";
            bunifuTextBox4.Text = "";
            bunifuDropdown1.SelectedItem = null;
            bunifuDropdown2.SelectedItem = null;
            bunifuDropdown3.SelectedItem = null;
            bunifuDropdown4.SelectedItem = null;
            checkBox1.Checked = false;
            bunifuCustomLabel1.Text = comboBox5.Items[comboBox5.SelectedIndex].ToString();
            Form2_Load(sender, e);
        }

        private void button_Click(object sender, EventArgs e)
        {

        }

        private void bunifuFlatButton1_Click(object sender, EventArgs e)
        {
            try
            {
                if (bunifuCustomLabel1.Text == "Заказчики")
                {
                    con.Open();
                    OleDbCommand cmd = con.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "INSERT INTO Заказчики (Фамилия, Имя, Отчество, Номер_Телефона) VALUES('" + bunifuTextBox1.Text + "','" + bunifuTextBox2.Text + "','" + bunifuTextBox3.Text + "','" + bunifuTextBox4.Text + "')";
                    cmd.ExecuteNonQuery();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = Queries.selectZakazch;
                    cmd.ExecuteNonQuery();
                    con.Close();
                    System.Data.DataTable dt = new System.Data.DataTable();
                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(dt);
                    bunifuDataGridView1.DataSource = dt;
                    MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    this.Refresh();
                }
                else
            if (bunifuCustomLabel1.Text == "Заказы")
                {
                    int client = getIdByClient(bunifuDropdown1.Text);
                    int Sotr = getIdBySotrudniki(bunifuDropdown2.Text);
                    con.Open();
                    OleDbCommand cmd = con.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "INSERT INTO Заказы (Id_Заказчика, Оформляющий, Адрес_Помещения, Сумма_Заказа, Дата_Заказа) VALUES(" + client + ", " + Sotr + ", '" + bunifuTextBox1.Text + "', '" + bunifuTextBox2.Text + "','" + bunifuDatePicker2.Value.ToString("dd.MM.yyyy") + "')";
                    cmd.ExecuteNonQuery();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = Queries.selectZakazi;
                    cmd.ExecuteNonQuery();
                    con.Close();
                    System.Data.DataTable dt = new System.Data.DataTable();
                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(dt);
                    bunifuDataGridView1.DataSource = dt;
                    MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
            if (bunifuCustomLabel1.Text == "Материалы")
                {

                    int Sklad = getIdBySklad(bunifuDropdown1.Text);
                    con.Open();
                    OleDbCommand cmd = con.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "INSERT INTO Материалы (Наименование, Цена_За_1, Колво, Id_Склада) VALUES('" + bunifuTextBox1.Text + "', " + bunifuTextBox2.Text + ", " + bunifuTextBox3.Text + ", " + Sklad + ")";
                    cmd.ExecuteNonQuery();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = Queries.selectMat;
                    cmd.ExecuteNonQuery();
                    con.Close();
                    System.Data.DataTable dt = new System.Data.DataTable();
                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(dt);
                    bunifuDataGridView1.DataSource = dt;
                    MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
            if (bunifuCustomLabel1.Text == "Работники")
                {
                    con.Open();
                    OleDbCommand cmd = con.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "INSERT INTO Работники (Профессия, Фамилия, Имя, Отчество, Номер_Телефона) VALUES('" + bunifuDropdown1.Text + "','" + bunifuTextBox1.Text + "', '" + bunifuTextBox2.Text + "','" + bunifuTextBox3.Text + "', '" + bunifuTextBox4.Text + "')";
                    cmd.ExecuteNonQuery();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = Queries.selectRabotniki;
                    cmd.ExecuteNonQuery();
                    con.Close();
                    System.Data.DataTable dt = new System.Data.DataTable();
                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(dt);
                    bunifuDataGridView1.DataSource = dt;
                    MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
            if (bunifuCustomLabel1.Text == "Поставщики")
                {
                    con.Open();
                    OleDbCommand cmd = con.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "INSERT INTO Поставщики (Название_организации, УНП, Юр_Адрес, Номер_Телефона) VALUES( '" + bunifuTextBox1.Text + "', '" + bunifuTextBox2.Text + "','" + bunifuTextBox3.Text + "','" + bunifuTextBox4.Text + "')";
                    cmd.ExecuteNonQuery();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = Queries.selectPostavch;
                    cmd.ExecuteNonQuery();
                    con.Close();
                    System.Data.DataTable dt = new System.Data.DataTable();
                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(dt);
                    bunifuDataGridView1.DataSource = dt;
                    MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
            if (bunifuCustomLabel1.Text == "Поставки")
                {
                    int postavshik = getIdByPostavshik(bunifuDropdown1.Text);
                    int material = getIdByMaterial(bunifuDropdown2.Text);
                    int sklad = getIdBySklad(bunifuDropdown3.Text);
                    int rab = getIdBySotrudniki(bunifuDropdown4.Text);
                    con.Open();
                    OleDbCommand cmd = con.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "INSERT INTO Поставки (Id_Поставщика, Id_Материала, Id_Склада, Дата_Разгрузки, ОТветственный_За_Поставку) VALUES( " + postavshik + "," + material + ", " + sklad + ", '" + bunifuDatePicker2.Value.ToString("dd.MM.yyyy") + "', " + rab + ")";
                    cmd.ExecuteNonQuery();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = Queries.selectPostavki;
                    cmd.ExecuteNonQuery();
                    con.Close();
                    System.Data.DataTable dt = new System.Data.DataTable();
                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(dt);
                    bunifuDataGridView1.DataSource = dt;
                    MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
            if (bunifuCustomLabel1.Text == "Склады")
                {

                    con.Open();
                    OleDbCommand cmd = con.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "INSERT INTO Склады ( Адрес_Склада) VALUES( '" + bunifuTextBox1.Text + "')";
                    cmd.ExecuteNonQuery();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = Queries.selectScladi;
                    cmd.ExecuteNonQuery();
                    con.Close();
                    System.Data.DataTable dt = new System.Data.DataTable();
                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(dt);
                    bunifuDataGridView1.DataSource = dt;
                    MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                else
                if (bunifuCustomLabel1.Text == "Пользователи")
                {
                    con.Open();
                    OleDbCommand cmd = con.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "INSERT INTO Пользователи (Логин,Пароль,Право_Администрации) VALUES( '" + bunifuTextBox1.Text + "', '" + bunifuTextBox2.Text + "', '" + bunifuDropdown1.Text + "')";
                    cmd.ExecuteNonQuery();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = Queries.selectPolz;
                    cmd.ExecuteNonQuery();
                    con.Close();
                    System.Data.DataTable dt = new System.Data.DataTable();
                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(dt);
                    bunifuDataGridView1.DataSource = dt;
                    MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);


                }
                bunifuTextBox1.Clear();
                bunifuTextBox2.Clear();
                bunifuTextBox3.Clear();
                bunifuTextBox4.Clear();
                bunifuDropdown1.Text = text1;
                bunifuDropdown2.Text = text2;
                bunifuDropdown3.Text = "";
                bunifuDropdown4.Text = "";
                comboBox5.Text = "";

            }
            catch (Exception)
            {

                 MessageBox.Show("Произошла ошибка, проверте корректность введённых данных", "Ошибка", MessageBoxButtons.OK);
            }
            


        }

        private void bunifuImageButton2_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void bunifuFlatButton2_Click(object sender, EventArgs e)
        {
            int ID = Convert.ToInt32(bunifuDataGridView1.CurrentRow.Cells[0].Value.ToString());

            if (bunifuCustomLabel1.Text == "Заказчики")
            {
                
                
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "UPDATE Заказчики SET Фамилия='" + bunifuTextBox1.Text + "', Имя='" + bunifuTextBox2.Text + "', Отчество='" + bunifuTextBox3.Text + "', Номер_Телефона='" + bunifuTextBox4.Text + "' WHERE Id_Заказчика= " + ID + "";
                con.Open();
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = Queries.selectZakazch;
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                bunifuDataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            if (bunifuCustomLabel1.Text == "Склады")
            {
                
                
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "UPDATE Склады SET Адрес_Склада='" + bunifuTextBox1.Text + "' WHERE id_Склада= " + ID + "";
                con.Open();
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = Queries.selectScladi;
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                bunifuDataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            if (bunifuCustomLabel1.Text == "Материалы")
            {                                                                     
                
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "UPDATE Материалы SET Наименование='" + bunifuTextBox1.Text + "',Цена_За_1= " + bunifuTextBox2.Text + ", Колво=" + bunifuTextBox3.Text + ", Id_Склада=" + getIdBySklad(bunifuDropdown1.Text) + " WHERE Код_Материала= " + ID + "";
                con.Open();
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = Queries.selectMat;
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                bunifuDataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            if (bunifuCustomLabel1.Text == "Заказы")
            {
                
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "UPDATE Заказы SET Id_Заказчика=" + getIdByClient(bunifuDropdown1.Text) + ", Оформляющий=" + getIdBySotrudniki(bunifuDropdown2.Text) + ", Адрес_Помещения='" + bunifuTextBox1.Text + "',Сумма_Заказа = " + bunifuTextBox2.Text + ", Дата_Заказа = '" + bunifuDatePicker2.Value.ToString("dd.MM.yyyy") + "' WHERE Id_Заказа= " + ID + "";
                con.Open();
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = Queries.selectZakazi;
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                bunifuDataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            
            if (bunifuCustomLabel1.Text == "Поставки")
            {
                
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "UPDATE Поставки SET Id_Поставщика=" + getIdByPostavshik(bunifuDropdown1.Text) + ", Id_Материала=" + getIdByMaterial(bunifuDropdown2.Text) + ", Id_Склада=" + getIdBySklad(bunifuDropdown3.Text) + ", Дата_Разгрузки='" + bunifuDatePicker2.Value.ToString("dd.MM.yyyy") + "', Ответственный_За_Поставку= " + getIdBySotrudniki(bunifuDropdown4.Text) + " WHERE Id_Поставки= " + ID + "";
                con.Open();
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = Queries.selectPostavki;
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                bunifuDataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            if (bunifuCustomLabel1.Text == "Поставщики")
            {
                
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "UPDATE Поставщики SET Название_организации='" + bunifuTextBox1.Text + "', УНП='" + bunifuTextBox2.Text + "', Юр_Адрес='" + bunifuTextBox3.Text + "', Номер_Телефона = '" + bunifuTextBox4.Text + "' WHERE Id_поставщика= " + ID + "";
                con.Open();
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = Queries.selectPostavch;
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                bunifuDataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            if (bunifuCustomLabel1.Text == "Работники")
            {
                
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "UPDATE Работники SET Профессия='" + bunifuTextBox1.Text + "', Фамилия='" + bunifuTextBox2.Text + "', Имя='" + bunifuTextBox3.Text + "', Отчество = '" + bunifuTextBox4.Text + "', Номер_Телефона = '" + bunifuTextBox6.Text + "' WHERE Id_Работника= " + ID + "";
                con.Open();
                cmd.ExecuteNonQuery();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = Queries.selectRabotniki;
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                bunifuDataGridView1.DataSource = dt;
                MessageBox.Show("Изменения сохранены!", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            bunifuTextBox1.Clear();
            bunifuTextBox2.Clear();
            bunifuTextBox3.Clear();
            bunifuTextBox4.Clear();
            bunifuDropdown3.Text = "";
            bunifuDropdown4.Text = "";
            comboBox5.Text = "";

        }

        

        private void dataGridView1_CellClick_1(object sender, DataGridViewCellEventArgs e)
        {
            
        }

        

        private void bunifuTextBox5_TextChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < bunifuDataGridView1.RowCount; i++)
            {
                for (int j = 0; j < bunifuDataGridView1.ColumnCount; j++)
                {
                    if (bunifuDataGridView1.Rows[i].Cells[j].Value != null)
                    {
                        bunifuDataGridView1.Rows[i].Cells[j].Value.ToString().ToLower();
                    }
                }
            }
            if (bunifuCustomLabel1.Text == "Заказы")
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT Заказы.Id_Заказа, Заказчики.Фамилия +' '+ Заказчики.Имя + ' ' + Заказчики.Отчество AS [ФИО Заказчика], Работники.Фамилия + ' ' + Работники.Имя + ' ' + Работники.Отчество AS [оформляющий], Заказы.Адрес_Помещения, Заказы.Сумма_Заказа, Заказы.Дата_Заказа FROM Работники INNER JOIN(Заказчики INNER JOIN Заказы ON Заказчики.Id_Заказчика = Заказы.Id_Заказчика) ON Работники.Id_Работника = Заказы.Оформляющий WHERE (Заказчики.Фамилия LIKE '%" + bunifuTextBox5.Text.ToLower() + "%'  or Заказчики.Имя LIKE '%" + bunifuTextBox5.Text.ToLower() + "%' or Заказчики.Отчество LIKE '%" + bunifuTextBox5.Text.ToLower() + "%' or Работники.Фамилия LIKE '%" + bunifuTextBox5.Text.ToLower() + "%' or Работники.Имя LIKE '%" + bunifuTextBox5.Text.ToLower() + "%' or Работники.Отчество LIKE '%" + bunifuTextBox5.Text.ToLower() + "%' or Адрес_Помещения LIKE '%" + bunifuTextBox5.Text.ToLower() + "%' or Сумма_Заказа LIKE '%" + bunifuTextBox5.Text.ToLower() + "%')";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                bunifuDataGridView1.DataSource = dt;
            }
            else
            if (bunifuCustomLabel1.Text == "Заказчики")
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM Заказчики WHERE (Заказчики.Фамилия LIKE '%" + bunifuTextBox5.Text.ToLower() + "%'  or Заказчики.Имя LIKE '%" + bunifuTextBox5.Text.ToLower() + "%' or Заказчики.Отчество LIKE '%" + bunifuTextBox5.Text.ToLower() + "%' or Номер_Телефона LIKE '%" + bunifuTextBox5.Text.ToLower() + "%') ";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                bunifuDataGridView1.DataSource = dt;
            }
            if (bunifuCustomLabel1.Text == "Материалы")
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT Материалы.Код_Материала, Материалы.Наименование, Материалы.Цена_За_1, Материалы.Колво, Склады.Адрес_Склада FROM Склады INNER JOIN Материалы ON Склады.id_Склада = Материалы.Id_Склада WHERE  (Материалы.Наименование LIKE '%" + bunifuTextBox5.Text.ToLower() + "%' or Материалы.Цена_За_1 LIKE '%" + bunifuTextBox5.Text.ToLower() + "%' or Материалы.Колво LIKE '%" + bunifuTextBox5.Text.ToLower() + "%' or Склады.Адрес_Склада LIKE '%" + bunifuTextBox5.Text.ToLower() + "%' ) ";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                bunifuDataGridView1.DataSource = dt;
            }
            if (bunifuCustomLabel1.Text == "Поставщики")
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM Поставщики WHERE  (Название_организации LIKE '%" + bunifuTextBox5.Text.ToLower() + "%' or УНП LIKE '%" + bunifuTextBox5.Text.ToLower() + "%' or Юр_Адрес LIKE '%" + bunifuTextBox5.Text.ToLower() + "%'or Номер_Телефона LIKE '%" + bunifuTextBox5.Text.ToLower() + "%')";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                bunifuDataGridView1.DataSource = dt;
            }
            if (bunifuCustomLabel1.Text == "Поставки")
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT Поставки.Id_Поставки, Поставщики.Название_организации, Материалы.Наименование, Склады.Адрес_Склада, Поставки.Дата_Разгрузки, Работники.Фамилия FROM Работники INNER JOIN(Поставщики INNER JOIN (Материалы INNER JOIN (Склады INNER JOIN Поставки ON Склады.id_Склада = Поставки.Id_Склада) ON(Склады.id_Склада = Материалы.Id_Склада) AND(Материалы.Код_Материала = Поставки.Id_Материала)) ON Поставщики.Id_поставщика = Поставки.Id_Поставщика) ON Работники.Id_Работника = Поставки.Ответсвенный_За_Поставку WHERE (Поставщики.Название_организации LIKE '%" + bunifuTextBox5.Text.ToLower() + "%' or Материалы.Наименование LIKE '%" + bunifuTextBox5.Text.ToLower() + "%' or Склады.Адрес_Склада LIKE '%" + bunifuTextBox5.Text.ToLower() + "%' or Поставки.Дата_Разгрузки LIKE '%" + bunifuTextBox5.Text.ToLower() + "%' or Работники.Фамилия LIKE '%" + bunifuTextBox5.Text.ToLower() + "%')";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                bunifuDataGridView1.DataSource = dt;
            }
            if (bunifuCustomLabel1.Text == "Склады")
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM Склады WHERE (Адрес_Склада LIKE '%" + bunifuTextBox5.Text.ToLower() + "%' )";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                bunifuDataGridView1.DataSource = dt;
            }
            if (bunifuCustomLabel1.Text == "Работники")
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM Работники WHERE(Работники.Фамилия LIKE '%" + bunifuTextBox5.Text.ToLower() + "%' or Работники.Имя LIKE '%" + bunifuTextBox5.Text.ToLower() + "%' or Работники.Отчество LIKE '%" + bunifuTextBox5.Text.ToLower() + "%' or Номер_Телефона LIKE '%" + bunifuTextBox5.Text.ToLower() + "%')";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                bunifuDataGridView1.DataSource = dt;
            }
           
        }

        
        public void Export_Data_To_Word(DataGridView DGV, string filename)
        {
            if (DGV.Rows.Count != 0)
            {
                int RowCount = DGV.Rows.Count;
                int ColumnCount = DGV.Columns.Count;
                Object[,] DataArray = new object[RowCount + 1, ColumnCount + 1];

                //add rows
                int r = 0;
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    for (r = 0; r <= RowCount - 1; r++)
                    {
                        DataArray[r, c] = DGV.Rows[r].Cells[c].Value;
                    } //end row loop
                } //end column loop

                Word.Document oDoc = new Word.Document();
                oDoc.Application.Visible = true;

                //page orintation
                oDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;


                dynamic oRange = oDoc.Content.Application.Selection.Range;
                string oTemp = "";
                for (r = 0; r <= RowCount - 1; r++)
                {
                    for (int c = 0; c <= ColumnCount - 1; c++)
                    {
                        oTemp = oTemp + DataArray[r, c] + "\t";

                    }
                }

                //table format
                oRange.Text = oTemp;

                object Separator = Word.WdTableFieldSeparator.wdSeparateByTabs;
                object ApplyBorders = true;
                object AutoFit = true;
                object AutoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitContent;

                oRange.ConvertToTable(ref Separator, ref RowCount, ref ColumnCount,
                Type.Missing, Type.Missing, ref ApplyBorders,
                Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, ref AutoFit, ref AutoFitBehavior, Type.Missing);

                oRange.Select();

                oDoc.Application.Selection.Tables[1].Select();
                oDoc.Application.Selection.Tables[1].Rows.AllowBreakAcrossPages = 0;
                oDoc.Application.Selection.Tables[1].Rows.Alignment = 0;
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.InsertRowsAbove(1);
                oDoc.Application.Selection.Tables[1].Rows[1].Select();

                //header row style
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Bold = 1;
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Name = "Tahoma";
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Size = 14;

                //add header row manually
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    oDoc.Application.Selection.Tables[1].Cell(1, c + 1).Range.Text = DGV.Columns[c].HeaderText;
                }

                //table style
                oDoc.Application.Selection.Tables[1].set_Style("Сетка таблицы");
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                //header text
                foreach (Word.Section section in oDoc.Application.ActiveDocument.Sections)
                {
                    Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                    headerRange.Text = bunifuCustomLabel1.Text;
                    headerRange.Font.Size = 16;
                    headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                }
                oDoc.SaveAs2(filename);
            }
        }
        private void bunifuFlatButton6_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();

            sfd.Filter = "Word Documents (*.docx)|*.docx";

            sfd.FileName = "export.docx";

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                Export_Data_To_Word(bunifuDataGridView1, sfd.FileName);
            }
        }
        private void QuitExcel()
        {
            if (this.xlWorkBook != null)
            {
                try
                {
                    this.xlWorkBook.Close();
                    Marshal.ReleaseComObject(this.xlWorkBook);
                }
                catch (COMException)
                {
                }

                this.xlWorkBook = null;
            }

            if (this.xlExcel != null)
            {
                try
                {
                    this.xlExcel.Quit();
                    Marshal.ReleaseComObject(this.xlExcel);
                }
                catch (COMException)
                {
                }

                this.xlExcel = null;
            }
        }
        private void CopyGrid()
        {
            // I'm making this up...
            bunifuDataGridView1.MultiSelect = true;
            bunifuDataGridView1.SelectAll();

            var data = bunifuDataGridView1.GetClipboardContent();

            if (data != null)
            {
                Clipboard.SetDataObject(data, true);
            }
            bunifuDataGridView1.MultiSelect =false;
        }
        private void bunifuFlatButton7_Click(object sender, EventArgs e)
        {
            try
            {
                this.QuitExcel();
                this.xlExcel = new Application { Visible = false };
                this.xlWorkBook = this.xlExcel.Workbooks.Add(Missing.Value);

                // Copy contents of grid into clipboard, open new instance of excel, a new workbook and sheet,
                // paste clipboard contents into new sheet.
                this.CopyGrid();

                var xlWorkSheet = (Worksheet)this.xlWorkBook.Worksheets.Item[1];

                try
                {
                    var cr = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[1, 1];

                    try
                    {
                        cr.Select();
                        xlWorkSheet.PasteSpecial(cr, NoHTMLFormatting: true);
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(cr);
                    }

                    this.xlWorkBook.SaveAs(Path.Combine(Path.GetTempPath(), "ItemUpdate.xls"), XlFileFormat.xlExcel5);
                }
                finally
                {
                    Marshal.ReleaseComObject(xlWorkSheet);
                }

                MessageBox.Show("File Save Successful", "Information", MessageBoxButtons.OK);

                //// If box is checked, show the exported file. Otherwise quit Excel.
                //if (this.checkBox1.Checked)
                //{
                this.xlExcel.Visible = true;
                //}
                //else
                //{
                // this.QuitExcel();
                //}
            }
            catch (SystemException ex)
            {
                MessageBox.Show(ex.ToString());
            }

            // Set the Selection Mode back to Cell Select to avoid conflict with sorting mode.
            this.bunifuDataGridView1.SelectionMode = DataGridViewSelectionMode.CellSelect;
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void bunifuFlatButton8_Click(object sender, EventArgs e)
        {
            bunifuGradientPanel2.Visible = true;
        }

        private void bunifuImageButton3_Click(object sender, EventArgs e)
        {
            string s = bunifuDatePicker1.Value.ToString("dd.MM.yyyy");
            string s2 = bunifuDatePicker3.Value.ToString("dd.MM.yyyy");
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT Заказы.Id_Заказа, Заказчики.Фамилия +' '+ Заказчики.Имя + ' ' + Заказчики.Отчество AS [ФИО Заказчика], Работники.Фамилия + ' ' + Работники.Имя + ' ' + Работники.Отчество AS [оформляющий], Заказы.Адрес_Помещения, Заказы.Сумма_Заказа, Заказы.Дата_Заказа FROM Работники INNER JOIN(Заказчики INNER JOIN Заказы ON Заказчики.Id_Заказчика = Заказы.Id_Заказчика) ON Работники.Id_Работника = Заказы.Оформляющий WHERE  Дата_Заказа >= @dateFirst and Дата_Заказа <= @dateSecond";
            cmd.Parameters.AddWithValue("@dateFirst", s);
            cmd.Parameters.AddWithValue("@dateSecond", s2);
            cmd.ExecuteNonQuery();
            con.Close();
            System.Data.DataTable dt = new System.Data.DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);
            bunifuDataGridView1.DataSource = dt;
            bunifuGradientPanel2.Visible = false;
        }

        private void bunifuPictureBox1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Данная автоматизированая систему была создала для того, чтобы упростить работу ИП Бужинский Е.И , которая помогает следить за оформлением заказов.. В программном продукте присутствуют все необходимые компоненты для ведения учёта заказов. Благодарим за пользование программой", "Уведомление", MessageBoxButtons.OK);
        }

        private void bunifuGradientPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void bunifuDataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            
        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void bunifuFlatButton5_Click(object sender, EventArgs e)
        {
            OleDbConnection con2 = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=E:\Практика\Tecn.accdb");
            Sostav f3 = new Sostav();
            f3.Text = "Заполнение заказа";
            int id_ZAK = int.Parse(bunifuDataGridView1.Rows[bunifuDataGridView1.CurrentRow.Index].Cells[0].Value.ToString());
            con2.Open();
            OleDbCommand cmd2 = con2.CreateCommand();
            cmd2.CommandType = CommandType.Text;
            cmd2.CommandText = "SELECT Состав_Заказа.Id_Состава_Заказа, Материалы.Наименование, Состав_Заказа.[Кол-во] FROM Материалы INNER JOIN Состав_Заказа ON Материалы.Код_Материала = Состав_Заказа.Код_Материала WHERE Id_Заказа = " + bunifuDataGridView1.Rows[bunifuDataGridView1.CurrentRow.Index].Cells[0].Value + "";
            cmd2.ExecuteNonQuery();
            System.Data.DataTable dt2 = new System.Data.DataTable();
            OleDbDataAdapter da2 = new OleDbDataAdapter(cmd2);
            da2.Fill(dt2);
            con2.Close();
            f3.bunifuDataGridView1.DataSource = dt2;
            f3.bunifuCustomLabel1.Text = "Заказ " + bunifuDataGridView1.Rows[bunifuDataGridView1.CurrentRow.Index].Cells[0].Value.ToString();
            f3.bunifuDataGridView1.Columns[0].Visible = false;
            f3.Show();
        }

        private void bunifuImageButton4_Click(object sender, EventArgs e)
        {
            if (bunifuDataGridView1.Height == 479)
            {
                bunifuDataGridView1.Height = bunifuDataGridView1.Height -168;
                panel1.Visible = true;
                bunifuDataGridView2.Visible = true;
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT Состав_Заказа.Id_Состава_Заказа, Заказы.Id_Заказа, Материалы.Наименование, Материалы.Колво FROM Материалы INNER JOIN(Заказы INNER JOIN Состав_Заказа ON Заказы.Id_Заказа = Состав_Заказа.Id_Заказа) ON Материалы.Код_Материала = Состав_Заказа.Код_Материала where Заказы.id_заказа =" + bunifuDataGridView1.Rows[bunifuDataGridView1.CurrentRow.Index].Cells[0].Value + "";
                cmd.ExecuteNonQuery();
                con.Close();
                System.Data.DataTable dt = new System.Data.DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                bunifuDataGridView2.DataSource = dt;
                bunifuDataGridView2.Columns[0].Visible = false;
            }
            else
            {
                bunifuDataGridView1.Height = bunifuDataGridView1.Height + 168;
                panel1.Visible =false;
                bunifuDataGridView2.Visible = false;
               
            }
            
        }

        private void bunifuDataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex != -1)
            {
                if (bunifuCustomLabel1.Text == "Материалы")
                {
                    bunifuTextBox1.Text = bunifuDataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                    bunifuTextBox2.Text = bunifuDataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
                    bunifuTextBox3.Text = bunifuDataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
                    bunifuDropdown1.Text = bunifuDataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
                }
                if (bunifuCustomLabel1.Text == "Склады")
                {
                    bunifuTextBox1.Text = bunifuDataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();


                }
                if (bunifuCustomLabel1.Text == "Заказы")
                {
                    bunifuTextBox1.Text = bunifuDataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
                    bunifuTextBox2.Text = bunifuDataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
                    bunifuDropdown1.Text = bunifuDataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                    bunifuDropdown2.Text = bunifuDataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
                    bunifuDatePicker2.Value = DateTime.Parse(bunifuDataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString());
                }
                if (bunifuCustomLabel1.Text == "Заказчики")
                {
                    bunifuTextBox1.Text = bunifuDataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                    bunifuTextBox2.Text = bunifuDataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
                    bunifuTextBox3.Text = bunifuDataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
                    bunifuTextBox4.Text = bunifuDataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();

                }

                if (bunifuCustomLabel1.Text == "Поставки")
                {
                    bunifuDropdown1.Text = bunifuDataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                    bunifuDropdown2.Text = bunifuDataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
                    bunifuDropdown3.Text = bunifuDataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
                    bunifuDatePicker2.Value = DateTime.Parse(bunifuDataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString());
                    bunifuDropdown4.Text = bunifuDataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString();

                }
                if (bunifuCustomLabel1.Text == "Поставщики")
                {
                    bunifuTextBox1.Text = bunifuDataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                    bunifuTextBox2.Text = bunifuDataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
                    bunifuTextBox3.Text = bunifuDataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
                    bunifuTextBox4.Text = bunifuDataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
                }
                if (bunifuCustomLabel1.Text == "Работники")
                {
                    bunifuTextBox1.Text = bunifuDataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                    bunifuTextBox2.Text = bunifuDataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
                    bunifuTextBox3.Text = bunifuDataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
                    bunifuTextBox4.Text = bunifuDataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
                    bunifuTextBox6.Text = bunifuDataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
                }


            }
        }

    }
}
