using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using static System.Net.Mime.MediaTypeNames;
using Microsoft.Office.Interop.Word;

namespace MyHotel
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }
        Form3 form3 = new Form3();
        private SqlConnection sqlConnection = null;
        private SqlCommand sqlCommand = null;
        private SqlDataReader sqlReader = null;
        string App_id; //id заявки, которая не принята
        
        private void Form2_Load(object sender, EventArgs e)
        {
            sqlConnection = new SqlConnection(@"Data Source=ERVOPC;Initial Catalog=MyHotel;Integrated Security=True");
            sqlConnection.Open();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string query = @"Select a.Application_id, a.Type as 'Тип заявки', a.Room_number as 'Номер комнаты','Подробнее' AS [Свойства] from Application as a where a.Execution_status = 'ожидание принятия'";
            Zayavka(query);
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            panel4.Visible = false;
            panel5.Location = new System.Drawing.Point(300, 94);
            panel5.Size = new Size(539, 559);
            panel4.Visible = false;
            panel7.Visible = false;
            panel5.Visible = true;
            try
            {
                string task = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
                if (task == "Подробнее")
                {
                    int indexrow = e.RowIndex;
                    App_id = dataGridView1.Rows[indexrow].Cells[0].Value.ToString();
                    string name = dataGridView1.Rows[indexrow].Cells[1].Value.ToString();
                    sqlCommand = new SqlCommand("SELECT * FROM Application", sqlConnection);
                    try
                    {
                        sqlReader = sqlCommand.ExecuteReader();
                        while (sqlReader.Read())
                        {
                            if (Convert.ToString(sqlReader["Application_id"]) == App_id && Convert.ToString(sqlReader["Type"]) == name)
                            {
                                textBox4.Text = Convert.ToString(sqlReader["Application_id"]);
                                textBox1.Text = Convert.ToString(sqlReader["Guest_name"]);
                                textBox2.Text = Convert.ToString(sqlReader["Room_number"]);
                                textBox3.Text = Convert.ToString(sqlReader["Type"]);
                                textBox5.Text = Convert.ToString(sqlReader["Submission_time"]);
                                textBox6.Text = Convert.ToString(sqlReader["Submission_date"]);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        sqlReader?.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            string query = @"Select a.Application_id, a.Type as 'Тип заявки', a.Room_number as 'Номер комнаты','Подробнее' AS [Свойства] from Application as a where a.Execution_status = 'ожидание принятия'";
            panel5.Visible = false;
            panel5.Location = new System.Drawing.Point(845, 94);
            panel5.Size = new Size(539, 559);
            panel6.Visible = false;
            panel7.Visible = false;
            panel4.Visible = true;
            Zayavka(query);
        }

        //Сотрудник берет заявку
        private void button4_Click(object sender, EventArgs e)
        {
            string query = @"Select a.Application_id, a.Type as 'Тип заявки', a.Room_number as 'Номер комнаты','Подробнее' AS [Свойства] from Application as a where a.Execution_status = 'ожидание принятия'";
            bool empapp = EmpApp();
            //id заявки высвечивавется
            if (empapp)
            {
                App_Update("Application","принято к выполнению");
                MessageBox.Show($"Заявка успешно зарегестрирована!");
                EmPApp_Ins();
                panel5.Visible = false;
                panel5.Location = new System.Drawing.Point(845, 94);
                panel5.Size = new Size(539, 559);
                panel4.Visible = true;
                Zayavka(query);
            }
            else
                MessageBox.Show("У вас уже есть активная заявка!");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string query = @"Select a.Application_id, a.Type as 'Тип заявки', a.Room_number as 'Номер комнаты','Подробнее' AS [Свойства] from Application as a where a.Execution_status = 'принято к выполнению'";
            Active_App(query);
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            panel4.Visible = false;
            panel6.Location = new System.Drawing.Point(300, 94);
            panel6.Size = new Size(539, 559);
            panel4.Visible = false;
            panel5.Visible = false;
            panel7.Visible = false;
            panel6.Visible = true;
            try
            {
                string task = dataGridView2.Rows[e.RowIndex].Cells[3].Value.ToString();
                if (task == "Подробнее")
                {
                    int indexrow = e.RowIndex;
                    App_id = dataGridView2.Rows[indexrow].Cells[0].Value.ToString();
                    string name = dataGridView2.Rows[indexrow].Cells[1].Value.ToString();
                    sqlCommand = new SqlCommand("SELECT * FROM Application", sqlConnection);
                    try
                    {
                        sqlReader = sqlCommand.ExecuteReader();
                        while (sqlReader.Read())
                        {
                            if (Convert.ToString(sqlReader["Application_id"]) == App_id && Convert.ToString(sqlReader["Type"]) == name)
                            {
                                textBox9.Text = Convert.ToString(sqlReader["Application_id"]);
                                textBox13.Text = label10.Text;
                                textBox12.Text = Convert.ToString(sqlReader["Guest_name"]);
                                textBox11.Text = Convert.ToString(sqlReader["Room_number"]);
                                textBox10.Text = Convert.ToString(sqlReader["Type"]);
                                textBox8.Text = Convert.ToString(sqlReader["Submission_time"]);
                                textBox7.Text = Convert.ToString(sqlReader["Submission_date"]);
                            }
                        }
                    }

                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    finally
                    {
                        sqlReader?.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string TemplateFileName = @"C:\Users\ErVo\Desktop\MyHotel\Blank_shab\First.docx";
        private void button6_Click(object sender, EventArgs e)
        {
            Form2 form2 = new Form2();
            string id = Convert.ToString(textBox9.Text);
            string guest_name = Convert.ToString(textBox12.Text);
            string emp_name = Convert.ToString(textBox13.Text);
            string room_numb = Convert.ToString(textBox11.Text);
            string type = Convert.ToString(textBox10.Text);
            string time = Convert.ToString(textBox8.Text);
            string date = Convert.ToString(textBox7.Text);

            var wordApp = new Word.Application();
            wordApp.Visible = false;

            var wordDocument = wordApp.Documents.Open(TemplateFileName);
            Word("{id}", id, wordDocument);
            Word("{guest_name}", guest_name, wordDocument);
            Word("{emp_name}", emp_name, wordDocument);
            Word("{room_numb}", room_numb, wordDocument);
            Word("{type}", type, wordDocument);
            Word("{time}", time, wordDocument);
            Word("{date}", date, wordDocument);
            Word("{emp1_name}", emp_name, wordDocument);
            Word("{date1}", date, wordDocument);

            wordApp.Visible = true;
        }


        private void pictureBox3_Click(object sender, EventArgs e)
        {
            string query = @"Select a.Application_id, a.Type as 'Тип заявки', a.Room_number as 'Номер комнаты','Подробнее' AS [Свойства] from Application as a where a.Execution_status = 'принято к выполнению'";
            panel5.Visible = false;
            panel5.Location = new System.Drawing.Point(845, 94);
            panel5.Size = new Size(539, 559);
            panel6.Visible = false;
            panel7.Visible = false;
            panel4.Visible = true;
            Active_App(query);
        }

        //завершение выполнения заявки
        private void button5_Click(object sender, EventArgs e)
        {
            DateTime now = DateTime.Now;//текущие дата и время
            string date = Convert.ToString(now.Date);
            string time = Convert.ToString(now.TimeOfDay);
            sqlCommand = new SqlCommand($@"UPDATE Application
                SET Execution_status = 'выполнено',
                Completion_date = '{date}',
                Completion_time = '{time}'
                Where Application_id = @App_id", sqlConnection);
            sqlCommand.Parameters.AddWithValue("@App_id", App_id);
            sqlCommand.ExecuteNonQuery();
            MessageBox.Show("Хорошая работа!");
            App_Update("EmpApp", "выполнено");
            panel6.Visible = false;
            string query = @"Select a.Application_id, a.Type as 'Тип заявки', a.Room_number as 'Номер комнаты','Подробнее' AS [Свойства] from Application as a where a.Execution_status = 'принято к выполнению'";
            Active_App(query);
        }


        //получить id Сотрудника
        public int Emp_id()
        {
            int emp_id = 0;
            string name = Convert.ToString(label10.Text);
            SqlCommand sqlCommand = new SqlCommand(@"Select * from Employee", sqlConnection);
            try
            {
                sqlReader = sqlCommand.ExecuteReader();
                while (sqlReader.Read())
                {
                    if (Convert.ToString(sqlReader["Full_name"]) == name)
                    {
                        emp_id = Convert.ToInt32(sqlReader["Employee_id"]);
                        return emp_id;
                    }
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            finally
            {
                sqlReader?.Close();
            }
            return emp_id;
        }

        private void Zayavka(string query)
        {
            dataGridView3.Visible = false;
            dataGridView2.Visible = false;
            dataGridView1.Visible= true;
            panel7.Visible = false;
            panel5.Visible = false;
            panel5.Location = new System.Drawing.Point(845, 94);
            panel5.Size = new Size(539, 559);
            label1.Text = "Доступные заявки";
            panel4.Visible = true;
            SqlCommand zayavka = new SqlCommand(query, sqlConnection);
            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(zayavka);
            System.Data.DataTable table = new System.Data.DataTable();
            sqlDataAdapter.Fill(table);
            dataGridView1.DataSource = table;
            dataGridView1.ReadOnly = true;
            DataGridViewCellStyle style = dataGridView1.ColumnHeadersDefaultCellStyle;
            style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            DataGridViewCellStyle style1 = dataGridView1.DefaultCellStyle;
            style1.Alignment = DataGridViewContentAlignment.MiddleCenter;
            this.dataGridView1.Columns["Application_id"].Visible = false;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                DataGridViewLinkCell linkCell = new DataGridViewLinkCell();
                dataGridView1[3, i] = linkCell;
            }
        }

        bool empapp;
        private bool EmpApp()
        {
            string status = "принято к выполнению";
            sqlCommand = new SqlCommand("SELECT * FROM Application", sqlConnection);
            try
            {
                sqlReader = sqlCommand.ExecuteReader();
                while (sqlReader.Read())
                {
                    if (Convert.ToString(sqlReader["Application_id"]) == Convert.ToString(App_id) && Convert.ToString(sqlReader["Execution_status"]) == status)
                        return empapp = false;

                    else
                        return empapp = true;
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            finally
            {
                if (sqlReader != null)
                    sqlReader.Close();
                sqlCommand = null;
            }
            return empapp;
        }

        private void App_Update(string table, string status)
        {
            sqlCommand = new SqlCommand($@"UPDATE {table}
                SET Execution_status = '{status}'
                Where Application_id = @App_id", sqlConnection);
            sqlCommand.Parameters.AddWithValue("@App_id", App_id);
            sqlCommand.ExecuteNonQuery();            
        }

        private void EmPApp_Ins()
        {
            sqlCommand = new SqlCommand(@"INSERT INTO EmpApp (Employee_id, Application_id,Employee_name,Execution_status) 
                Values (@Employee_id, @Application_id, @Employee_name, @Execution_status)", sqlConnection);
            sqlCommand.Parameters.AddWithValue("Employee_id", Emp_id());
            sqlCommand.Parameters.AddWithValue("Application_id", App_id);
            sqlCommand.Parameters.AddWithValue("Employee_name", label10.Text);
            sqlCommand.Parameters.AddWithValue("Execution_status", "принято к выполнению");
            sqlCommand.ExecuteNonQuery();
        }

        private void Active_App(string query)
        {
            dataGridView3.Visible = false;
            dataGridView1.Visible = false;
            dataGridView2.Visible = true;
            panel7.Visible = false;
            panel5.Visible = false;
            panel5.Location = new System.Drawing.Point(845, 95);
            panel5.Size = new Size(539, 559);
            label1.Text = "Активные заявки";
            panel4.Visible = true;
            SqlCommand zayavka = new SqlCommand(query, sqlConnection);
            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(zayavka);
            System.Data.DataTable table = new System.Data.DataTable();
            sqlDataAdapter.Fill(table);
            dataGridView2.DataSource = table;
            dataGridView2.ReadOnly = true;
            DataGridViewCellStyle style = dataGridView2.ColumnHeadersDefaultCellStyle;
            style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            DataGridViewCellStyle style1 = dataGridView2.DefaultCellStyle;
            style1.Alignment = DataGridViewContentAlignment.MiddleCenter;
            this.dataGridView2.Columns["Application_id"].Visible = false;
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                DataGridViewLinkCell linkCell = new DataGridViewLinkCell();
                dataGridView2[3, i] = linkCell;
            }
        }

        private void Word(string stub, string text, Word.Document wordDocument)
        {
            var range = wordDocument.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: stub, ReplaceWith: text);
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Form1 form = new Form1();
            form.Visible = true;
            this.Close();
        }

        
        private void Completed_App(string query)
        {
            dataGridView1.Visible = false;
            dataGridView2.Visible = false;
            dataGridView3.Visible = true;
            panel7.Visible = false;
            panel5.Visible = false;
            panel5.Location = new System.Drawing.Point(845, 95);
            panel5.Size = new Size(539, 559);
            label1.Text = "Выполненные заявки";
            panel4.Visible = true;
            SqlCommand zayavka = new SqlCommand(query, sqlConnection);
            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(zayavka);
            System.Data.DataTable table = new System.Data.DataTable();
            sqlDataAdapter.Fill(table);
            dataGridView3.DataSource = table;
            dataGridView3.ReadOnly = true;
            DataGridViewCellStyle style = dataGridView3.ColumnHeadersDefaultCellStyle;
            style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            DataGridViewCellStyle style1 = dataGridView3.DefaultCellStyle;
            style1.Alignment = DataGridViewContentAlignment.MiddleCenter;
            this.dataGridView3.Columns["Application_id"].Visible = false;
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                DataGridViewLinkCell linkCell = new DataGridViewLinkCell();
                dataGridView3[3, i] = linkCell;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string query = @"Select a.Application_id, a.Type as 'Тип заявки', a.Room_number as 'Номер комнаты','Подробнее' AS [Свойства] from Application as a where a.Execution_status = 'выполнено'";
            Completed_App(query);
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            string query = @"Select a.Application_id, a.Type as 'Тип заявки', a.Room_number as 'Номер комнаты','Подробнее' AS [Свойства] from Application as a where a.Execution_status = 'выполнено'";
            panel7.Visible = false;
            panel7.Location = new System.Drawing.Point(845, 95);
            panel7.Size = new Size(539, 559);
            label1.Text = "Выполненные заявки";
            panel4.Visible = true;
            Completed_App(query);
        }

        private void dataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            panel4.Visible = false;
            panel7.Location = new System.Drawing.Point(300, 94);
            panel7.Size = new Size(539, 559);
            panel4.Visible = false;
            panel5.Visible = false;
            panel6.Visible = true;
            panel7.Visible = true;
            try
            {
                string task = dataGridView3.Rows[e.RowIndex].Cells[3].Value.ToString();
                if (task == "Подробнее")
                {
                    int indexrow = e.RowIndex;
                    App_id = dataGridView3.Rows[indexrow].Cells[0].Value.ToString();
                    string name = dataGridView3.Rows[indexrow].Cells[1].Value.ToString();
                    sqlCommand = new SqlCommand("SELECT * FROM Application", sqlConnection);
                    try
                    {
                        sqlReader = sqlCommand.ExecuteReader();
                        while (sqlReader.Read())
                        {
                            if (Convert.ToString(sqlReader["Application_id"]) == App_id && Convert.ToString(sqlReader["Type"]) == name)
                            {
                                textBox17.Text = Convert.ToString(sqlReader["Application_id"]);
                                textBox20.Text = label10.Text;
                                textBox14.Text = Convert.ToString(sqlReader["Guest_name"]);
                                textBox19.Text = Convert.ToString(sqlReader["Room_number"]);
                                textBox18.Text = Convert.ToString(sqlReader["Type"]);
                                textBox16.Text = Convert.ToString(sqlReader["Submission_time"]);
                                textBox15.Text = Convert.ToString(sqlReader["Submission_date"]);
                                textBox22.Text = Convert.ToString(sqlReader["Completion_time"]);
                                textBox21.Text = Convert.ToString(sqlReader["Completion_date"]);
                            }
                        }
                    }

                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    finally
                    {
                        sqlReader?.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string TemplateFileName1 = @"C:\Users\ErVo\Desktop\MyHotel\Blank_shab\Second.docx";
        private void button7_Click(object sender, EventArgs e)
        {
            Form2 form2 = new Form2();
            string id = Convert.ToString(textBox17.Text);
            string guest_name = Convert.ToString(textBox14.Text);
            string emp_name = Convert.ToString(textBox20.Text);
            string room_numb = Convert.ToString(textBox19.Text);
            string type = Convert.ToString(textBox18.Text);
            string time = Convert.ToString(textBox16.Text);
            string date = Convert.ToString(textBox15.Text);
            string after_date = Convert.ToString(textBox21.Text);
            string after_time = Convert.ToString(textBox22.Text);

            var wordApp = new Word.Application();
            wordApp.Visible = false;

            var wordDocument = wordApp.Documents.Open(TemplateFileName1);
            Word("{id}", id, wordDocument);
            Word("{guest_name}", guest_name, wordDocument);
            Word("{emp_name}", emp_name, wordDocument);
            Word("{room_numb}", room_numb, wordDocument);
            Word("{type}", type, wordDocument);
            Word("{time}", time, wordDocument);
            Word("{date}", date, wordDocument);
            Word("{time1}", after_time, wordDocument);
            Word("{date2}", after_date, wordDocument);
            Word("{emp1_name}", emp_name, wordDocument);
            Word("{date1}", date, wordDocument);
            Word("{guest1_name}", guest_name, wordDocument);
            Word("{date3}", date, wordDocument);

            wordApp.Visible = true;
        }
    }
}
