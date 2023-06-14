using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Reflection.Emit;

namespace MyHotel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private SqlConnection sqlConnection = null;
        private SqlCommandBuilder sqlBuilder = null;
        private SqlDataAdapter sqlDataAdapter = null;
        private DataSet dataSet = null;
        private SqlCommand sqlCommand = null;
        private SqlDataReader sqlReader = null;

        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "myHotelDataSet.Service". При необходимости она может быть перемещена или удалена.
            this.serviceTableAdapter.Fill(this.myHotelDataSet.Service);
            sqlConnection = new SqlConnection(@"Data Source=ERVOPC;Initial Catalog=MyHotel;Integrated Security=True");
            sqlConnection.Open();
            textBox9.PasswordChar = '*';
            //string abobus = "ПАШО НАХУЙ СА СВАИМ ПРАГРАМИРОВАНИЕМ";
            //string huita = $@"Хуй тебе а не пятерку и да, {abobus}";
            //MessageBox.Show(huita);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            label15.Visible = false;
            panel4.Visible = false;
            panel6.Location = new Point(300, 95);
            panel6.Size = new Size(539,556);
            panel7.Visible = false;
            panel7.Location = new Point(845, 95);
            panel7.Size = new Size(539, 559);
            panel6.Visible = true;
        }

        //заполнить датагридвью таблицей Уcлуги
        private void Uslugi()
        {
            try
            {
                sqlDataAdapter = new SqlDataAdapter("SELECT Service_id, Service_name, 'Подробнее' AS [More] FROM Service", sqlConnection);
                sqlBuilder = new SqlCommandBuilder(sqlDataAdapter);
                
                dataSet = new DataSet();
                sqlDataAdapter.Fill(dataSet, "Service");
                dataGridView1.DataSource = dataSet.Tables["Service"];
                this.dataGridView1.Columns["Service_id"].Visible = false;
                dataGridView1.ReadOnly = true;
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();
                    dataGridView1[2, i] = linkCell;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            label15.Visible = false;
            panel6.Location = new Point(845, 95);
            panel6.Size = new Size(539, 559);
            panel6.Visible = false;
            panel7.Visible = false;
            panel7.Location = new Point(845, 95);
            panel7.Size = new Size(539, 559);
            panel5.Visible = false;
            panel5.Location = new Point(845, 95);
            panel5.Size = new Size(539, 559);
            panel4.Visible = true;
            Uslugi();
        }

        //добавить заявку
        private void button4_Click(object sender, EventArgs e)
        {
            string name = textBox1.Text;
            int roomNumb = Convert.ToInt32(textBox2.Text);
            string type = comboBox1.Text;
            int guest_id = 0;
            string status = Convert.ToString("ожидание принятия");

            DateTime now = DateTime.Now;//текущие дата и время
            string date = Convert.ToString(now.Date);
            string time = Convert.ToString(now.TimeOfDay);
            
            try
            {
                    sqlCommand = new SqlCommand("SELECT * FROM GUEST", sqlConnection);
                    try
                    {
                        sqlReader = sqlCommand.ExecuteReader();
                        while (sqlReader.Read())
                        {                        
                            if (Convert.ToString(sqlReader["Name"]) == name)
                            {
                              guest_id = Convert.ToInt32(sqlReader["Guest_id"]);
                              continue;
                            }
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
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            sqlCommand = new SqlCommand("INSERT INTO Application (Guest_name, Room_number, Type, Submission_date, Submission_time," +
                "Execution_status,Guest_id)VALUES(@Guest_name, @Room_number, @Type, @Submission_date, @Submission_time, " +
                "@Execution_status,@Guest_id)", sqlConnection);

            if (!string.IsNullOrEmpty(name) && !string.IsNullOrWhiteSpace(Convert.ToString(roomNumb)) && !string.IsNullOrEmpty(type) )
            {
                sqlCommand.Parameters.AddWithValue("Guest_name", name);
                sqlCommand.Parameters.AddWithValue("Room_number", roomNumb);
                sqlCommand.Parameters.AddWithValue("Type", type);
                sqlCommand.Parameters.AddWithValue("Submission_date",date);
                sqlCommand.Parameters.AddWithValue("Submission_time", time);
                sqlCommand.Parameters.AddWithValue("Execution_status", status);
                sqlCommand.Parameters.AddWithValue("Guest_id", guest_id);
                sqlCommand.ExecuteNonQuery();
                MessageBox.Show("Заявка добавлена в список. Ожидайте");

                textBox1.Text = null;
                textBox2.Text = null;
                comboBox1.Text = null;
            }
            else
            {
               MessageBox.Show("Поля должны быть заполнены.");
            }
        }

        //получить "Подробнее" о заявке
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            panel4.Visible=false;            
            panel6.Visible=false;
            panel5.Location = new Point(300, 95);
            panel5.Size = new Size(539, 556);
            panel7.Visible = false;
            panel7.Location = new Point(845, 95);
            panel7.Size = new Size(539, 559);
            panel5.Visible = true;

            try
            {
                string task = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
                if (task == "Подробнее")
                {
                  int indexrow = e.RowIndex;
                  string id = dataGridView1.Rows[indexrow].Cells[0].Value.ToString();
                  string name = dataGridView1.Rows[indexrow].Cells[1].Value.ToString();
                  sqlCommand = new SqlCommand("SELECT * FROM SERVICE", sqlConnection);
                  try
                   {
                     sqlReader = sqlCommand.ExecuteReader();
                     while (sqlReader.Read())
                     {
                       if (Convert.ToString(sqlReader["Service_id"]) == id && Convert.ToString(sqlReader["Service_name"]) == name)
                       {
                         textBox3.Text = Convert.ToString(sqlReader["Service_name"]);
                         richTextBox1.Text = Convert.ToString(sqlReader["Description"]);
                         textBox4.Text = Convert.ToString(sqlReader["Price"] + " руб.");
                         textBox5.Text = Convert.ToString(sqlReader["Realisation_time"] + " мин.");
                       }
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
                  }
                }

                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        

        private void button5_Click(object sender, EventArgs e)
        {
            label15.Visible = false;
            panel5.Visible = false;
            panel5.Location = new Point(845, 95);
            panel5.Size = new Size(539, 559);
            panel6.Visible = false;
            panel6.Location = new Point(845, 95);
            panel6.Size = new Size(539, 559);
            panel7.Visible = false;
            panel7.Location = new Point(845, 95);
            panel7.Size = new Size(539, 559);
            textBox6.ReadOnly= true;
            textBox7.ReadOnly= true;
            panel4.Visible = true;
        }

        //проверка статуса заявки
        private void button6_Click(object sender, EventArgs e)
        {
            textBox6.Clear();
            sqlCommand = new SqlCommand("SELECT * FROM Application", sqlConnection);

            try
            {
                sqlReader = sqlCommand.ExecuteReader();

                while (sqlReader.Read())
                {
                    if (Convert.ToString(sqlReader["Guest_name"]) == Convert.ToString(textBox7.Text) &&
                        Convert.ToString(sqlReader["Type"]) == Convert.ToString(comboBox2.Text))
                    {
                        textBox6.Text = Convert.ToString(sqlReader["Execution_status"]);
                    }
                }
                if(string.IsNullOrEmpty(textBox6.Text))
                MessageBox.Show("Заявка не найдена");
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

        private void button1_Click(object sender, EventArgs e)
        {
            label15.Visible = false;
            panel5.Visible = false;
            panel5.Location = new Point(845, 95);
            panel5.Size = new Size(539, 559);
            panel6.Visible = false;
            panel6.Location = new Point(845, 95);
            panel6.Size = new Size(539, 559);
            panel4.Visible = false;
            panel7.Location = new Point(300, 95);
            panel7.Size = new Size(539, 556);
            panel7.Visible = true;
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            label15.Visible = false;
            linkLabel1.Visible= false;
            panel1.Visible= false;
            panel5.Visible = false;
            panel5.Location = new Point(845, 95);
            panel5.Size = new Size(539, 559);
            panel6.Visible = false;
            panel6.Location = new Point(845, 95);
            panel6.Size = new Size(539, 559);
            panel4.Visible = false;
            panel7.Location = new Point(845, 95);
            panel7.Size = new Size(539, 556);
            panel7.Visible = false;
            panel8.Location = new Point(0, 96);
            panel8.Size = new Size(839, 557);
            panel8.Visible = true;
            panel2.Visible = false; 
        }

        
        Form2 form2 = new Form2();
        Form3 form3 = new Form3();

        //авторизация
        private void button7_Click(object sender, EventArgs e)
        {
            string name;
            sqlCommand = new SqlCommand("SELECT * FROM Employee", sqlConnection);

            try
            {
                sqlReader = sqlCommand.ExecuteReader();

                while (sqlReader.Read())
                {
                    if (Convert.ToString(sqlReader["Login"]) == Convert.ToString(textBox8.Text) && Convert.ToString(sqlReader["Password"]) == Convert.ToString(textBox9.Text))
                    {
                       
                        if (Convert.ToString(sqlReader["Post"]) == Convert.ToString("Сотрудник"))
                        {
                            name = Convert.ToString(sqlReader["Full_name"]);
                            form2.label10.Text = name;
                            this.Visible=false;
                            form2.Show();                                                        
                        }
                        else if (Convert.ToString(sqlReader["Post"]) == Convert.ToString("Администратор"))
                        {
                            name = Convert.ToString(sqlReader["Full_name"]);
                            form3.Show();
                            form3.label10.Text = name;
                            this.Visible = false;
                        }

                    }
                    else
                    { label15.Text = "Пароль или логин введены неверно."; label15.Visible = true; }

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
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == false)
            {
                textBox9.PasswordChar = '*';
                //textBox5.PasswordChar = '*';
            }

            else
            {
                textBox9.PasswordChar = '\0';
                //textBox5.PasswordChar = '\0';
            }
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            label15.Visible = false;
            linkLabel1.Visible = true;
            panel1.Visible = true;
            panel5.Visible = false;
            panel5.Location = new Point(845, 95);
            panel5.Size = new Size(539, 559);
            panel6.Visible = false;
            panel6.Location = new Point(845, 95);
            panel6.Size = new Size(539, 559);
            panel7.Location = new Point(845, 95);
            panel7.Size = new Size(539, 556);
            panel7.Visible = false;
            panel8.Location = new Point(845, 95);
            panel8.Size = new Size(839, 557);
            panel8.Visible = false;
            panel2.Visible = true;
            panel4.Visible = true;
        }
    }
}
