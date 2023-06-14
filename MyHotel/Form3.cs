using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Net.NetworkInformation;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace MyHotel
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }

        private SqlConnection sqlConnection = null;
        private SqlCommand sqlCommand = null;
        private SqlDataReader sqlReader = null;
        string organisation1 = @"Select * from Organization";
        string emoployee1 = @"Select * from Employee";
        string room1 = @"Select * from Room";
        string guest1 = @"Select * from Guest";
        string application1 = @"Select * from Application";
        string service1 = @"Select * from Service";
        string organisation = @"Select 
                o.Organisation_id as 'Номер',
                o.INN as 'ИНН',
                o.OGRN as 'ОГРН',
                o.Org_name as 'Организация',
                o.Adress as 'Адрес',
                o.Director_name as 'Директор',
                o.Employees_count as 'Число сотрудников',
                o.Room_count as 'Число комнат'
                from Organization as o";
        string employee = @"Select 
                e.Employee_id as 'Номер',
                e.Full_name as 'Полное имя',
                e.Passport_series as 'Серия паспорта',
                e.Passport_number as 'Номер паспорта',
                e.Post as 'Должность',
                e.Phone_number as 'Номер телефона',
                e.Organisation_id as 'Номер отдела',
                e.Login as 'Логин',
                e.Password as 'Пароль'
                from Employee as e";
        string room = @"Select 
                r.Room_number as 'Номер комнаты',
                r.Room_class as 'Класс комнаты',
                r.Amount_sleeping_place as 'Количество спальных мест',
                r.Amount_begs as 'Количество кроватей',
                r.Child_beg as 'Количество детских кроватей',
                r.Beg_type as 'Тип кроватей',
                r.Price as 'Цена за сутки',
                r.Organisation_id as 'Номер отдела организации'
                from Room as r";
        string guest = @"Select
                g.Guest_id as 'Номер',
                g.Name as 'Полное имя',
                g.Passport_series as 'Серия паспорта',
                g.Passport_number as 'Номер паспорта',
                g.Phone_number as 'Номер телефона',
                g.Room_number as 'Номер комнаты',
                g.Settlement_date as 'Дата заселения',
                g.Settlement_time as 'Время заселения',
                g.Eviction_date as 'Дата выселения',
                g.Eviction_time as 'Время выселения'
                from Guest as g";
        string application = @"Select
                a.Application_id as 'Номер',
                a.Guest_name as 'Полное имя',
                a.Room_number as 'Серия паспорта',
                a.Type as 'Номер паспорта',
                a.Submission_date as 'Номер телефона',
                a.Submission_time as 'Номер комнаты',
                a.Execution_status as 'Дата заселения',
                a.Guest_id as 'Время заселения',
                a.Completion_date as 'Дата выселения',
                a.Completion_time as 'Время выселения'
                from Application as a";
        string service = @"Select
                s.Service_id as 'Номер заявки',
                s.Service_name as 'Название заявки',
                s.Description as 'Описание',
                s.Price as 'Цена',
                s.Realisation_time as 'Время реализации'
                From Service as s";
        string empapp = @"Select
                ea.Employee_id as 'Номер сотрудника',
                ea.Application_id as 'Номер заявки',
                ea.Employee_name as 'Имя сотрудника',
                ea.Execution_status as 'Статус'
                From EmpApp as ea";
        string roomserv = @"Select
                rs.Room_number as 'Номер комнаты',
                rs.Service_id as 'Номер услуги',
                rs.Service_name as 'Называние услуги'
                From RoomServ as rs";
        string delete_org = @"Delete from Organization where Organisation_id = @text";
        string delete_emp = @"Delete from Employee where Employee_id = @text";
        string delete_room = @"Delete from Room where Room_number = @text";
        string delete_guest = @"Delete from Guest where Guest_id = @text";
        string delete_application = @"Delete from Application where Application_id = @text";
        string delete_service = @"Delete from Service where Service_id = @text";

        private void Form3_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "myHotelDataSet7.Service". При необходимости она может быть перемещена или удалена.
            this.serviceTableAdapter1.Fill(this.myHotelDataSet7.Service);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "myHotelDataSet6.Service". При необходимости она может быть перемещена или удалена.
            this.serviceTableAdapter.Fill(this.myHotelDataSet6.Service);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "myHotelDataSet5.Application". При необходимости она может быть перемещена или удалена.
            this.applicationTableAdapter.Fill(this.myHotelDataSet5.Application);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "myHotelDataSet4.Guest". При необходимости она может быть перемещена или удалена.
            this.guestTableAdapter.Fill(this.myHotelDataSet4.Guest);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "myHotelDataSet3.Room". При необходимости она может быть перемещена или удалена.
            this.roomTableAdapter.Fill(this.myHotelDataSet3.Room);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "myHotelDataSet2.Employee". При необходимости она может быть перемещена или удалена.
            this.employeeTableAdapter.Fill(this.myHotelDataSet2.Employee);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "myHotelDataSet1.Organization". При необходимости она может быть перемещена или удалена.
            this.organizationTableAdapter.Fill(this.myHotelDataSet1.Organization);
            sqlConnection = new SqlConnection(@"Data Source=ERVOPC;Initial Catalog=MyHotel;Integrated Security=True");
            sqlConnection.Open();
            textBox11.MaxLength = 4; //серия паспорта
            textBox10.MaxLength = 6; //номер паспорта
            textBox23.MaxLength = 4; //серия паспорта
            textBox22.MaxLength = 6; //номер паспорта
            textBox42.MaxLength = 4; //серия паспорта
            textBox41.MaxLength = 6; //номер паспорта
            textBox54.MaxLength = 4; //серия паспорта
            textBox53.MaxLength = 6; //номер паспорта
            textBox16.MaxLength = 11; //номер телефона
            textBox17.MaxLength = 11; //номер телефона
            textBox40.MaxLength = 11; //номер телефона
            textBox52.MaxLength = 11; //номер телефона
            textBox56.ReadOnly = true;
            textBox55.ReadOnly = true;
            textBox43.ReadOnly= true;
            textBox59.ReadOnly = true;
            textBox68.ReadOnly = true;
            Methods.Reload(organisation, dataGridView1);
            Methods.Reload(employee, dataGridView2);
            Methods.Reload(room, dataGridView3);
            Methods.Reload(guest, dataGridView4);
            Methods.Reload(application, dataGridView5);
            Methods.Reload(service, dataGridView6);
            Methods.Reload(empapp, dataGridView7);
            Methods.Reload(roomserv, dataGridView8);
        }
            
        private void radioButton1_CheckedChanged(object sender, EventArgs e) // панельдобавления для орагнизации
        {
            panel3.Visible = false;
            panel2.Visible = false;
            panel1.Visible = true;
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e) //панель редактирвоания для организации
        {
            panel4.Visible = false;
            panel1.Visible = false;
            panel3.Visible = true;
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e) // панель удалени для организации
        {
            panel3.Visible = false;
            panel1.Visible = false;
            panel4.Visible = true;
        }

        private void button1_Click(object sender, EventArgs e) //обновить организацию
        {
            Methods.Reload(organisation, dataGridView1);
        }

        private void button4_Click(object sender, EventArgs e) //удаление организации
        {
            int id = Convert.ToInt32(comboBox1.Text);
            Methods.Delete(delete_org, id);
            Methods.Reload(organisation, dataGridView1);
            this.organizationTableAdapter.Fill(this.myHotelDataSet1.Organization);
        }
                                
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e) //комбобокс редактирования
        {
            string text = comboBox1.Text;
            sqlCommand = new SqlCommand(organisation1, sqlConnection);
            try
            {
                sqlReader = sqlCommand.ExecuteReader();
                while (sqlReader.Read())
                {
                    if (Convert.ToString(sqlReader["Organisation_id"]) == text)
                    {
                        textBox8.Text = Convert.ToString(sqlReader["Adress"]);
                        textBox7.Text = Convert.ToString(sqlReader["Director_name"]);
                        textBox6.Text = Convert.ToString(sqlReader["Employees_count"]);
                        textBox5.Text = Convert.ToString(sqlReader["Room_count"]);
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

        private void button3_Click(object sender, EventArgs e) //кнопка редактирования
        {
            string text = comboBox1.Text;            
            bool org = Methods.Check(organisation1, text, "Organisation_id");
            if (org == true)
            {
                if (!string.IsNullOrEmpty(textBox8.Text) && !string.IsNullOrEmpty(textBox7.Text) && !string.IsNullOrEmpty(textBox6.Text) && !string.IsNullOrEmpty(textBox5.Text) && !string.IsNullOrEmpty(comboBox1.Text))
                {
                    if (int.TryParse(comboBox1.Text, out int id) == true)
                    {
                        if (int.TryParse(textBox6.Text, out int Employees_count) == true && int.TryParse(textBox5.Text, out int Room_count) == true)
                        {
                            try
                            {
                                sqlCommand = new SqlCommand(@"UPDATE Organization 
                                SET Adress = @Adress,
                                    Director_name = @Director_name,
                                    Employees_count = @Employees_count,
                                    Room_count = @Room_count
                                WHERE Organisation_id = @Organisation_id", sqlConnection);
                                sqlCommand.Parameters.AddWithValue("@Adress", Convert.ToString(textBox8.Text));
                                sqlCommand.Parameters.AddWithValue("@Director_name", Convert.ToString(textBox7.Text));
                                sqlCommand.Parameters.AddWithValue("@Employees_count", Convert.ToInt32(textBox6.Text));
                                sqlCommand.Parameters.AddWithValue("@Room_count", Convert.ToInt32(textBox5.Text));
                                sqlCommand.Parameters.AddWithValue("@Organisation_id", Convert.ToInt32(comboBox1.Text));
                                sqlCommand.ExecuteNonQuery();
                                MessageBox.Show("Редактирование прошло успешно!");
                                Methods.Reload(organisation, dataGridView1);
                            }
                            catch (Exception exp)
                            {
                                MessageBox.Show("Что-то пошло не так.\n" + exp);
                            }
                        }
                        else { MessageBox.Show("Поля 'Количество сотрудников' и/или 'Количество комнат' имеют неверный формат значений"); }
                    }
                    MessageBox.Show("Номер отдела введен некорректно");
                }
                else
                    MessageBox.Show("Поля должны быть заполнены!");
            }
            else
                MessageBox.Show("Такой орагнизации нет, проверьте поле 'Номер'!");
        }

        //добавление
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (!string.IsNullOrEmpty(textBox1.Text) && !string.IsNullOrEmpty(textBox2.Text) && !string.IsNullOrEmpty(textBox3.Text) && !string.IsNullOrEmpty(textBox4.Text))
                {
                    if (int.TryParse(textBox3.Text, out int Employees_count) == true && int.TryParse(textBox4.Text, out int Room_count) == true)
                    {
                        try
                        {
                            sqlCommand = new SqlCommand(@"INSERT INTO Organization (INN,OGRN,Org_name,Adress,Director_name,Employees_count,Room_count) 
                                                        Values(31212345345,432534523453,'Отель MyHotel',@Adress,@Director_name,@Employees_count,@Room_count)", sqlConnection);
                            sqlCommand.Parameters.AddWithValue("@Adress", Convert.ToString(textBox1.Text));
                            sqlCommand.Parameters.AddWithValue("@Director_name", Convert.ToString(textBox2.Text));
                            sqlCommand.Parameters.AddWithValue("@Employees_count", Convert.ToInt32(textBox3.Text));
                            sqlCommand.Parameters.AddWithValue("@Room_count", Convert.ToInt32(textBox4.Text));
                            sqlCommand.ExecuteNonQuery();
                            MessageBox.Show("Запись успечшно добавлена!");
                            Methods.Reload(organisation, dataGridView1);
                            this.organizationTableAdapter.Fill(this.myHotelDataSet1.Organization);
                        }
                        catch (Exception exp)
                        {
                            MessageBox.Show("Что-то пошло не так.\n" + exp);
                        }
                    }
                    else { MessageBox.Show("Поля 'Количество сотрудников' и/или 'Количество комнат' имеют неверный формат значений"); }
                }
                else
                    MessageBox.Show("Поля должны быть заполнены!");
            }
            catch
            {
                MessageBox.Show("Что-то пошло не так! Повторите попытку снова");
            }
        }

        private void radioButton6_CheckedChanged(object sender, EventArgs e) //панель добавление сотрудника
        {
            panel5.Visible = false;
            panel6.Visible = false;
            panel7.Visible= true;
        }

        private void radioButton5_CheckedChanged(object sender, EventArgs e) //панель редакктирования пользователя
        {
            panel5.Visible = false;
            panel7.Visible = false;
            panel6.Visible = true;
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e) //панел удаления пользователя
        {
            panel7.Visible = false;
            panel6.Visible = false;
            panel5.Visible = true;
        }

        private void button8_Click(object sender, EventArgs e) //обновить сотрудников
        {
            Methods.Reload(employee, dataGridView2);
        }

        private void button5_Click(object sender, EventArgs e) //удаление сотрудника
        {
            int id = Convert.ToInt32(comboBox3.Text);
            Methods.Delete(delete_emp, id);
            Methods.Reload(employee, dataGridView2);
            this.employeeTableAdapter.Fill(this.myHotelDataSet2.Employee);
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e) // комббокс редактирования сотрдуника
        {
            string text = comboBox4.Text;
            sqlCommand = new SqlCommand(emoployee1, sqlConnection);
            try
            {
                sqlReader = sqlCommand.ExecuteReader();
                while (sqlReader.Read())
                {
                    if (Convert.ToString(sqlReader["Employee_id"]) == text)
                    {
                        textBox12.Text = Convert.ToString(sqlReader["Full_name"]);
                        textBox11.Text = Convert.ToString(sqlReader["Passport_series"]);
                        textBox10.Text = Convert.ToString(sqlReader["Passport_number"]);
                        textBox9.Text = Convert.ToString(sqlReader["Post"]);
                        textBox17.Text = Convert.ToString(sqlReader["Phone_number"]);
                        textBox18.Text = Convert.ToString(sqlReader["Organisation_id"]);
                        textBox19.Text = Convert.ToString(sqlReader["Login"]);
                        textBox20.Text = Convert.ToString(sqlReader["Password"]);
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

        private void button6_Click(object sender, EventArgs e) //кнопка редактирование сотрудника
        {
            string text = comboBox4.Text;
            bool emp = Methods.Check(emoployee1, text, "Employee_id");
            string text1 = textBox18.Text;
            bool org = Methods.Check(organisation1, text1, "Organisation_id");
            bool phone = Methods.PhoneCheked(textBox17.Text);

            if (phone == true)
            {
                if (emp == true)
                {
                    if (!string.IsNullOrEmpty(comboBox4.Text) && !string.IsNullOrEmpty(textBox12.Text) && !string.IsNullOrEmpty(textBox11.Text)
                        && !string.IsNullOrEmpty(textBox10.Text) && !string.IsNullOrEmpty(textBox9.Text) && !string.IsNullOrEmpty(textBox17.Text)
                        && !string.IsNullOrEmpty(textBox18.Text) && !string.IsNullOrEmpty(textBox19.Text) && !string.IsNullOrEmpty(textBox20.Text))
                    {
                        if (int.TryParse(comboBox4.Text, out int id) == true)
                        {
                            if (int.TryParse(textBox11.Text, out int Passport_series) == true && int.TryParse(textBox10.Text, out int Passport_number) == true)
                            {
                                if (org == true)
                                {
                                    if (textBox9.Text == "Сотрудник" || textBox9.Text == "Администратор")
                                    {
                                        try
                                        {
                                            sqlCommand = new SqlCommand(@"UPDATE Employee 
                                            SET Full_name = @Full_name,
                                            Passport_series = @Passport_series,
                                            Passport_number = @Passport_number,
                                            Post = @Post,
                                            Phone_number = @Phone_number,
                                            Organisation_id = @Organisation_id,
                                            Login = @Login,
                                            Password = @Password
                                            WHERE Employee_id = @Employee_id", sqlConnection);
                                            sqlCommand.Parameters.AddWithValue("@Full_name", Convert.ToString(textBox12.Text));
                                            sqlCommand.Parameters.AddWithValue("@Passport_series", Convert.ToInt32(textBox11.Text));
                                            sqlCommand.Parameters.AddWithValue("@Passport_number", Convert.ToInt32(textBox10.Text));
                                            sqlCommand.Parameters.AddWithValue("@Post", Convert.ToString(textBox9.Text));
                                            sqlCommand.Parameters.AddWithValue("@Phone_number", Convert.ToString(textBox17.Text));
                                            sqlCommand.Parameters.AddWithValue("@Organisation_id", Convert.ToInt32(textBox18.Text));
                                            sqlCommand.Parameters.AddWithValue("@Login", Convert.ToString(textBox19.Text));
                                            sqlCommand.Parameters.AddWithValue("@Password", Convert.ToString(textBox20.Text));
                                            sqlCommand.Parameters.AddWithValue("@Employee_id", Convert.ToInt32(comboBox4.Text));
                                            sqlCommand.ExecuteNonQuery();
                                            MessageBox.Show("Редактирование прошло успешно!");
                                            Methods.Reload(employee, dataGridView2);
                                        }
                                        catch (Exception exp)
                                        {
                                            MessageBox.Show("Что-то пошло не так.\n" + exp);
                                        }
                                    }
                                    else
                                        MessageBox.Show("Должность введена неверно!");
                                }
                                else
                                    MessageBox.Show("Введен номер несуществующего отдела!");
                            }
                            else
                                MessageBox.Show("Поля 'Серия паспорта' и/или 'Номер паспорта' имеют неверный формат значений");
                        }
                        else
                        MessageBox.Show("Номер сотрудника введен некорректно");
                    }
                    else
                        MessageBox.Show("Поля должны быть заполнены!");
                }
                else
                    MessageBox.Show("Такой орагнизации нет, проверьте поле 'Номер'!"); 
            }
            else
            MessageBox.Show("Телефонный номер введен некорректно");
        }

        private void button7_Click(object sender, EventArgs e) //кнопка добавления сотрудника
        {
            string text1 = textBox15.Text;
            bool org = Methods.Check(organisation1, text1, "Organisation_id");
            bool phone = Methods.PhoneCheked(textBox16.Text);

            if (phone == true)
            {               
                if (!string.IsNullOrEmpty(textBox24.Text) && !string.IsNullOrEmpty(textBox23.Text)
                    && !string.IsNullOrEmpty(textBox22.Text) && !string.IsNullOrEmpty(textBox21.Text) && !string.IsNullOrEmpty(textBox16.Text)
                    && !string.IsNullOrEmpty(textBox15.Text) && !string.IsNullOrEmpty(textBox14.Text) && !string.IsNullOrEmpty(textBox13.Text))
                {
                        if (int.TryParse(textBox23.Text, out int Passport_series) == true && int.TryParse(textBox22.Text, out int Passport_number) == true)
                        {
                            if (org == true)
                            {
                                if (textBox21.Text == "Сотрудник" || textBox21.Text == "Администратор")
                                {
                                    try
                                    {
                                        sqlCommand = new SqlCommand(@"INSERT INTO Employee (Full_name, Passport_series, Passport_number,Post,
                                            Phone_number, Organisation_id, Login, Password)
                                            VALUES (@Full_name, @Passport_series, @Passport_number, @Post,
                                            @Phone_number, @Organisation_id, @Login, @Password)", sqlConnection);
                                        sqlCommand.Parameters.AddWithValue("@Full_name", Convert.ToString(textBox24.Text));
                                        sqlCommand.Parameters.AddWithValue("@Passport_series", Convert.ToInt32(textBox23.Text));
                                        sqlCommand.Parameters.AddWithValue("@Passport_number", Convert.ToInt32(textBox22.Text));
                                        sqlCommand.Parameters.AddWithValue("@Post", Convert.ToString(textBox21.Text));
                                        sqlCommand.Parameters.AddWithValue("@Phone_number", Convert.ToString(textBox16.Text));
                                        sqlCommand.Parameters.AddWithValue("@Organisation_id", Convert.ToInt32(textBox15.Text));
                                        sqlCommand.Parameters.AddWithValue("@Login", Convert.ToString(textBox14.Text));
                                        sqlCommand.Parameters.AddWithValue("@Password", Convert.ToString(textBox13.Text));
                                        sqlCommand.ExecuteNonQuery();
                                        this.employeeTableAdapter.Fill(this.myHotelDataSet2.Employee);
                                        MessageBox.Show("Сотрудник успешно добавлен!");
                                        Methods.Reload(employee, dataGridView2);
                                    }
                                    catch (Exception exp)
                                    {
                                        MessageBox.Show("Что-то пошло не так.\n" + exp);
                                    }
                                }
                                else
                                    MessageBox.Show("Должность введена неверно!");
                            }
                            else
                                MessageBox.Show("Введен номер несуществующего отдела!");
                        }
                        else
                            MessageBox.Show("Поля 'Серия паспорта' и/или 'Номер паспорта' имеют неверный формат значений");
                }
                else
                    MessageBox.Show("Поля должны быть заполнены!");                
            }
            else
                MessageBox.Show("Телефонный номер введен некорректно");            
        }

        private void radioButton9_CheckedChanged(object sender, EventArgs e)
        {
            panel8.Visible = false;
            panel9.Visible = false;
            panel10.Visible = true;
        }

        private void radioButton8_CheckedChanged(object sender, EventArgs e)
        {
            panel8.Visible = false;
            panel10.Visible = false;
            panel9.Visible = true;
        }

        private void radioButton7_CheckedChanged(object sender, EventArgs e)
        {
            panel9.Visible = false;
            panel10.Visible = false;
            panel8.Visible = true;
        }

        private void button12_Click(object sender, EventArgs e) // обновление для комнаты
        {
            Methods.Reload(room, dataGridView3);
        }

        private void button9_Click(object sender, EventArgs e) // удаление для комнаты
        {
            int id = Convert.ToInt32(comboBox5.Text);
            Methods.Delete(delete_room, id);
            Methods.Reload(room, dataGridView3);
            this.roomTableAdapter.Fill(this.myHotelDataSet3.Room);
        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            string text = comboBox6.Text;
            sqlCommand = new SqlCommand(room1, sqlConnection);
            try
            {
                sqlReader = sqlCommand.ExecuteReader();
                while (sqlReader.Read())
                {
                    if (Convert.ToString(sqlReader["Room_number"]) == text)
                    {
                        comboBox7.Text = Convert.ToString(sqlReader["Room_class"]);
                        textBox31.Text = Convert.ToString(sqlReader["Amount_sleeping_place"]);
                        textBox30.Text = Convert.ToString(sqlReader["Amount_begs"]);
                        textBox29.Text = Convert.ToString(sqlReader["Child_beg"]);
                        textBox28.Text = Convert.ToString(sqlReader["Beg_type"]);
                        textBox27.Text = Convert.ToString(sqlReader["Price"]);
                        textBox25.Text = Convert.ToString(sqlReader["Organisation_id"]);
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

        private void button10_Click(object sender, EventArgs e) //кнопка редактирования комнаты
        {
            string text = comboBox6.Text;
            bool rooms = Methods.Check(room1, text, "Room_number");
            string text1 = textBox25.Text;
            bool org = Methods.Check(organisation1, text1, "Organisation_id");            
            if (rooms == true)
            {
                if (!string.IsNullOrEmpty(comboBox6.Text) && !string.IsNullOrEmpty(comboBox7.Text) && !string.IsNullOrEmpty(textBox31.Text)
                    && !string.IsNullOrEmpty(textBox30.Text) && !string.IsNullOrEmpty(textBox29.Text) && !string.IsNullOrEmpty(textBox28.Text)
                    && !string.IsNullOrEmpty(textBox27.Text) && !string.IsNullOrEmpty(textBox25.Text))
                {
                    if (int.TryParse(comboBox6.Text, out int id) == true)
                    {
                        if (int.TryParse(textBox31.Text, out int Amount_sleeping_place) == true && int.TryParse(textBox30.Text, out int Amount_begs) == true
                            && int.TryParse(textBox29.Text, out int Child_beg) == true && double.TryParse(textBox27.Text, out double Price))
                        {
                            bool amount = Methods.Amount_sleep(Amount_sleeping_place, Amount_begs, Child_beg);
                            if (amount == true)
                            {                                
                                if (org == true)
                                {
                                    if (comboBox7.Text == "Эконом" || comboBox7.Text == "Средний класс" || comboBox7.Text == "Люкс")
                                    {
                                        try
                                        {
                                            sqlCommand = new SqlCommand(@"UPDATE Room 
                                                SET Room_class = @Room_class,
                                                Amount_sleeping_place = @Amount_sleeping_place,
                                                Amount_begs = @Amount_begs,
                                                Child_beg = @Child_beg,
                                                Beg_type = @Beg_type,
                                                Price = @Price,
                                                Organisation_id = @Organisation_id
                                                WHERE Room_number = @Room_number", sqlConnection);
                                            sqlCommand.Parameters.AddWithValue("@Room_class", Convert.ToString(comboBox7.Text));
                                            sqlCommand.Parameters.AddWithValue("@Amount_sleeping_place", Convert.ToInt32(textBox31.Text));
                                            sqlCommand.Parameters.AddWithValue("@Amount_begs", Convert.ToInt32(textBox30.Text));
                                            sqlCommand.Parameters.AddWithValue("@Child_beg", Convert.ToString(textBox29.Text));
                                            sqlCommand.Parameters.AddWithValue("@Beg_type", Convert.ToString(textBox28.Text));
                                            sqlCommand.Parameters.AddWithValue("@Price", Convert.ToDouble(textBox27.Text));
                                            sqlCommand.Parameters.AddWithValue("@Organisation_id", Convert.ToString(textBox25.Text));
                                            sqlCommand.Parameters.AddWithValue("@Room_number", Convert.ToInt32(comboBox6.Text));
                                            sqlCommand.ExecuteNonQuery();
                                            MessageBox.Show("Редактирование прошло успешно!");
                                            Methods.Reload(room, dataGridView3);
                                        }
                                        catch (Exception exp)
                                        {
                                            MessageBox.Show("Что-то пошло не так.\n" + exp);
                                        }
                                    }
                                    else
                                        MessageBox.Show("Наименование класса введено неверно!");
                                }
                                else
                                    MessageBox.Show("Введен номер несуществующего отдела!");
                            }
                            else
                                MessageBox.Show("Проверьте поля 'Количество спальных мест', \n 'Количество кроватей' и 'Количество детских кроватей'! \n " +
                                    "Необходимо, чтобы сумма значений последних двух была равна значению первого!");
                        }
                         else
                             MessageBox.Show("Поля 'Количество спальных мест' и/или 'Количество кроватей' \n" +
                                 "и/или 'Количество детских кроватей'и/или 'Цена за сутки' имеют неверный формат значений");
                    }
                    else
                        MessageBox.Show("Номер комнаты введен некорректно");
                }
                else
                    MessageBox.Show("Поля должны быть заполнены!");
            }
             else
                 MessageBox.Show("Такой комнаты нет, проверьте поле 'Номер'!");
            
        }

        private void button11_Click(object sender, EventArgs e) //добавление комнаты
        {
            string text1 = textBox26.Text;
            bool org = Methods.Check(organisation1, text1, "Organisation_id");
            
                if (!string.IsNullOrEmpty(comboBox8.Text) && !string.IsNullOrEmpty(textBox36.Text)
                    && !string.IsNullOrEmpty(textBox35.Text) && !string.IsNullOrEmpty(textBox34.Text) && !string.IsNullOrEmpty(textBox33.Text)
                    && !string.IsNullOrEmpty(textBox32.Text) && !string.IsNullOrEmpty(textBox26.Text))
                {
                    if (int.TryParse(textBox36.Text, out int Amount_sleeping_place) == true && int.TryParse(textBox35.Text, out int Amount_begs) == true
                            && int.TryParse(textBox34.Text, out int Child_beg) == true && double.TryParse(textBox32.Text, out double Price))
                    {
                        bool amount = Methods.Amount_sleep(Amount_sleeping_place, Amount_begs, Child_beg);
                        if (amount == true)
                        {
                            if (org == true)
                            {
                                if (comboBox8.Text == "Эконом" || comboBox8.Text == "Средний класс" || comboBox8.Text == "Люкс")
                                {
                                    try
                                    {
                                        sqlCommand = new SqlCommand(@"INSERT INTO Room (Room_class, Amount_sleeping_place, Amount_begs,
                                                Child_beg, Beg_type, Price, Organisation_id) 
                                                VALUES(@Room_class, @Amount_sleeping_place, @Amount_begs,
                                                @Child_beg, @Beg_type, @Price, @Organisation_id)", sqlConnection);
                                        sqlCommand.Parameters.AddWithValue("@Room_class", Convert.ToString(comboBox8.Text));
                                        sqlCommand.Parameters.AddWithValue("@Amount_sleeping_place", Convert.ToInt32(textBox36.Text));
                                        sqlCommand.Parameters.AddWithValue("@Amount_begs", Convert.ToInt32(textBox35.Text));
                                        sqlCommand.Parameters.AddWithValue("@Child_beg", Convert.ToString(textBox34.Text));
                                        sqlCommand.Parameters.AddWithValue("@Beg_type", Convert.ToString(textBox33.Text));
                                        sqlCommand.Parameters.AddWithValue("@Price", Convert.ToDouble(textBox32.Text));
                                        sqlCommand.Parameters.AddWithValue("@Organisation_id", Convert.ToString(textBox26.Text));
                                        sqlCommand.ExecuteNonQuery();
                                        this.roomTableAdapter.Fill(this.myHotelDataSet3.Room);
                                        MessageBox.Show("Комната добавлена успешно!");
                                        Methods.Reload(room, dataGridView3);
                                    }
                                    catch (Exception exp)
                                    {
                                        MessageBox.Show("Что-то пошло не так.\n" + exp);
                                    }
                                }
                                else
                                    MessageBox.Show("Наименование класса введено неверно!");
                            }
                            else
                                MessageBox.Show("Введен номер несуществующего отдела!");
                        }
                        else
                            MessageBox.Show("Проверьте поля 'Количество спальных мест', \n 'Количество кроватей' и 'Количество детских кроватей'! \n " +
                                "Необходимо, чтобы сумма значений последних двух была равна значению первого!");
                    }
                    else
                        MessageBox.Show("Поля 'Количество спальных мест' и/или 'Количество кроватей' \n" +
                            "и/или 'Количество детских кроватей'и/или 'Цена за сутки' имеют неверный формат значений");
                }
                else
                    MessageBox.Show("Поля должны быть заполнены!");
            
        }

        private void radioButton12_CheckedChanged(object sender, EventArgs e)
        {
            panel11.Visible = false;
            panel12.Visible = false;
            panel13.Visible = true;
        }

        private void radioButton11_CheckedChanged(object sender, EventArgs e)
        {
            panel11.Visible = false;
            panel13.Visible = false;
            panel12.Visible = true;
        }

        private void radioButton10_CheckedChanged(object sender, EventArgs e)
        {
            panel12.Visible = false;
            panel13.Visible = false;
            panel11.Visible = true;
        }

        private void button16_Click(object sender, EventArgs e) //обновление таблицы Гости
        {
            Methods.Reload(guest, dataGridView4);
        }

        private void button13_Click(object sender, EventArgs e)
        {
            int id = Convert.ToInt32(comboBox9.Text);
            Methods.Delete(delete_guest, id);
            Methods.Reload(guest, dataGridView4);
            this.guestTableAdapter.Fill(this.myHotelDataSet4.Guest);
        }

        private void comboBox11_SelectedIndexChanged(object sender, EventArgs e) //комбобокс таблицы Гости
        {
            string text = comboBox11.Text;
            sqlCommand = new SqlCommand(guest1, sqlConnection);
            try
            {
                sqlReader = sqlCommand.ExecuteReader();
                while (sqlReader.Read())
                {
                    if (Convert.ToString(sqlReader["Guest_id"]) == text)
                    {
                        textBox49.Text = Convert.ToString(sqlReader["Name"]);
                        textBox42.Text = Convert.ToString(sqlReader["Passport_series"]);
                        textBox41.Text = Convert.ToString(sqlReader["Passport_number"]);
                        textBox40.Text = Convert.ToString(sqlReader["Phone_number"]);
                        textBox39.Text = Convert.ToString(sqlReader["Room_number"]);
                        textBox38.Text = Convert.ToString(sqlReader["Settlement_date"]);
                        textBox37.Text = Convert.ToString(sqlReader["Settlement_time"]);
                        textBox51.Text = Convert.ToString(sqlReader["Eviction_date"]);
                        textBox50.Text = Convert.ToString(sqlReader["Eviction_time"]);
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

        private void button14_Click(object sender, EventArgs e) //редактирование таблицы гости
        {
            string text = comboBox11.Text;
            bool guest3 = Methods.Check(guest1, text, "Guest_id");            
            bool phone = Methods.PhoneCheked(textBox40.Text);
            if (guest3 == true)
            {
                if (!string.IsNullOrEmpty(comboBox11.Text) && !string.IsNullOrEmpty(textBox49.Text) && !string.IsNullOrEmpty(textBox41.Text)
                    && !string.IsNullOrEmpty(textBox40.Text) && !string.IsNullOrEmpty(textBox39.Text) && !string.IsNullOrEmpty(textBox38.Text)
                    && !string.IsNullOrEmpty(textBox37.Text))
                {
                    if (int.TryParse(comboBox11.Text, out int id) == true)
                    {
                        if (int.TryParse(textBox42.Text, out int Passport_series) == true && int.TryParse(textBox41.Text, out int Passport_number) == true)
                        {
                            if (phone == true)
                            {
                                if(Methods.DateChek(textBox38.Text) == true)
                                {
                                    if (Methods.TimeCheck(textBox37.Text) == true)
                                    {
                                        if (!string.IsNullOrEmpty(textBox51.Text) && !string.IsNullOrEmpty(textBox51.Text))
                                        {
                                            //выполняется редактирование с проверкой даты и времени
                                            if (Methods.DateChek(textBox51.Text) == true)
                                            {
                                                if (Methods.TimeCheck(textBox37.Text) == true)
                                                {
                                                    sqlCommand = new SqlCommand(@"UPDATE Guest 
                                                        SET Name = @Name,
                                                        Passport_series = @Passport_series,
                                                        Passport_number = @Passport_number,
                                                        Phone_number = @Phone_number,
                                                        Room_number = @Room_number,
                                                        Settlement_date = @Settlement_date,
                                                        Settlement_time = @Settlement_time,
                                                        Eviction_date = @Eviction_date,
                                                        Eviction_time = @Eviction_time
                                                        WHERE Guest_id = @Guest_id", sqlConnection);
                                                    sqlCommand.Parameters.AddWithValue("@Name", Convert.ToString(textBox49.Text));
                                                    sqlCommand.Parameters.AddWithValue("Passport_series", Convert.ToInt32(textBox42.Text));
                                                    sqlCommand.Parameters.AddWithValue("@Passport_number", Convert.ToInt32(textBox41.Text));
                                                    sqlCommand.Parameters.AddWithValue("@Phone_number", Convert.ToString(textBox40.Text));
                                                    sqlCommand.Parameters.AddWithValue("@Room_number", Convert.ToInt32(textBox39.Text));
                                                    sqlCommand.Parameters.AddWithValue("@Settlement_date", DateTime.Parse(textBox38.Text));
                                                    sqlCommand.Parameters.AddWithValue("@Settlement_time", DateTime.Parse(textBox37.Text));
                                                    sqlCommand.Parameters.AddWithValue("@Eviction_date", DateTime.Parse(textBox51.Text));
                                                    sqlCommand.Parameters.AddWithValue("@Eviction_time", DateTime.Parse(textBox50.Text));
                                                    sqlCommand.Parameters.AddWithValue("@Guest_id", Convert.ToInt32(comboBox11.Text));
                                                    sqlCommand.ExecuteNonQuery();
                                                    MessageBox.Show("Редактирование прошло успешно!");
                                                    Methods.Reload(guest, dataGridView4);
                                                }
                                                else
                                                    MessageBox.Show("Время выселения имеет неверный форат. Введите дату в формате 'чч:мм'");
                                            }
                                            else
                                                MessageBox.Show("Дата выселения имеет неверный формат. Введите дату в формате 'дд.мм.гггг'");
                                        }
                                        else
                                        {
                                            sqlCommand = new SqlCommand(@"UPDATE Guest 
                                                        SET Name = @Name,
                                                        Passport_series = @Passport_series,
                                                        Passport_number = @Passport_number,
                                                        Phone_number = @Phone_number,
                                                        Room_number = @Room_number,
                                                        Settlement_date = @Settlement_date,
                                                        Settlement_time = @Settlement_time                                                        
                                                        WHERE Guest_id = @Guest_id", sqlConnection);
                                            sqlCommand.Parameters.AddWithValue("@Name", Convert.ToString(textBox49.Text));
                                            sqlCommand.Parameters.AddWithValue("Passport_series", Convert.ToInt32(textBox42.Text));
                                            sqlCommand.Parameters.AddWithValue("@Passport_number", Convert.ToInt32(textBox41.Text));
                                            sqlCommand.Parameters.AddWithValue("@Phone_number", Convert.ToString(textBox40.Text));
                                            sqlCommand.Parameters.AddWithValue("@Room_number", Convert.ToInt32(textBox39.Text));
                                            sqlCommand.Parameters.AddWithValue("@Settlement_date", DateTime.Parse(textBox38.Text));
                                            sqlCommand.Parameters.AddWithValue("@Settlement_time", DateTime.Parse(textBox37.Text));                                            
                                            sqlCommand.Parameters.AddWithValue("@Guest_id", Convert.ToInt32(comboBox11.Text));
                                            sqlCommand.ExecuteNonQuery();
                                            MessageBox.Show("Редактирование прошло успешно!");
                                            Methods.Reload(guest, dataGridView4);
                                        }
                                    }
                                    else
                                        MessageBox.Show("Время заселения имеет неверный форат. Введите дату в формате 'чч:мм'");
                                }
                                else
                                    MessageBox.Show("Дата заселения имеет неверный формат. Введите дату в формате 'дд.мм.гггг'");
                            }
                            else
                                MessageBox.Show("Телефонный номер введен некорректно");
                        }
                        else
                            MessageBox.Show("Поля 'Серия паспорта' и/или 'Номер паспорта' имеют неверный формат значений");
                    }
                    else
                        MessageBox.Show("Номер гостя введен некорректно!");
                }
                else
                    MessageBox.Show("Поля должны быть заполнены!");
            }
            else
                MessageBox.Show("Такой гостя нет, проверьте поле 'Номер'!");
        }

        private void button15_Click(object sender, EventArgs e) //добавление посетителя
        {
            string text = textBox48.Text;
            bool room = Methods.Check(room1, text, "Room_number");
            bool phone = Methods.PhoneCheked(textBox52.Text);
            if (!string.IsNullOrEmpty(textBox45.Text) && !string.IsNullOrEmpty(textBox54.Text) && !string.IsNullOrEmpty(textBox53.Text) 
                && !string.IsNullOrEmpty(textBox52.Text) && !string.IsNullOrEmpty(textBox48.Text) && !string.IsNullOrEmpty(textBox47.Text) && !string.IsNullOrEmpty(textBox46.Text))
            {
                if (int.TryParse(textBox54.Text, out int Passport_series) == true && int.TryParse(textBox53.Text, out int Passport_number) == true)
                {
                    if (phone == true)
                    {
                        if (room == true)
                        {
                            if(Methods.DateChek(textBox47.Text) == true)
                            {
                                if (Methods.TimeCheck(textBox46.Text) == true)
                                {
                                    sqlCommand = new SqlCommand(@"INSERT INTO Guest (Name, Passport_series, Passport_number, Phone_number, Room_number,Settlement_date, Settlement_time)
                                        Values (@Name, @Passport_series, @Passport_number, @Phone_number, @Room_number, @Settlement_date, @Settlement_time)", sqlConnection);
                                    sqlCommand.Parameters.AddWithValue("@Name", Convert.ToString(textBox45.Text));
                                    sqlCommand.Parameters.AddWithValue("Passport_series", Convert.ToInt32(textBox54.Text));
                                    sqlCommand.Parameters.AddWithValue("@Passport_number", Convert.ToInt32(textBox53.Text));
                                    sqlCommand.Parameters.AddWithValue("@Phone_number", Convert.ToString(textBox52.Text));
                                    sqlCommand.Parameters.AddWithValue("@Room_number", Convert.ToInt32(textBox48.Text));
                                    sqlCommand.Parameters.AddWithValue("@Settlement_date", DateTime.Parse(textBox47.Text));
                                    sqlCommand.Parameters.AddWithValue("@Settlement_time", DateTime.Parse(textBox46.Text));
                                    sqlCommand.ExecuteNonQuery();
                                    MessageBox.Show("Добавление гостя прошло успешно!");
                                    this.guestTableAdapter.Fill(this.myHotelDataSet4.Guest);
                                    Methods.Reload(guest, dataGridView4);
                                }
                                else
                                    MessageBox.Show("Время заселения имеет неверный форат. Введите дату в формате 'чч:мм'");
                            }
                            else
                                MessageBox.Show("Дата заселения имеет неверный формат. Введите дату в формате 'дд.мм.гггг'");
                        }
                        else
                            MessageBox.Show("Выбранной комнаты не существует");
                    }
                    else
                        MessageBox.Show("Телефонный номер введен некорректно");
                }
                else
                    MessageBox.Show("Поля 'Серия паспорта' и/или 'Номер паспорта' имеют неверный формат значений");
            }
            else
                MessageBox.Show("Поля должны быть заполнены!");
        }

        private void radioButton15_CheckedChanged(object sender, EventArgs e) //панель добавления услуги
        {
            panel14.Visible = false;
            panel15.Visible= false;
            panel16.Visible= true;
        }

        private void radioButton14_CheckedChanged(object sender, EventArgs e) //панель редактирования услуги
        {
            panel14.Visible = false;
            panel16.Visible= false;
            panel15.Visible = true;
        }

        private void radioButton13_CheckedChanged(object sender, EventArgs e) //панель удаления услуги
        {
            panel15.Visible= false;
            panel16.Visible= false;
            panel14.Visible= true;
        }

        private void button20_Click(object sender, EventArgs e)
        {
            Methods.Reload(application, dataGridView5);
        }

        private void button17_Click(object sender, EventArgs e) //удаление заявки
        {
            int id = Convert.ToInt32(comboBox10.Text);
            Methods.Delete(delete_application, id);
            Methods.Reload(application, dataGridView5);
            this.applicationTableAdapter.Fill(this.myHotelDataSet5.Application);
        }

        private void comboBox13_SelectedIndexChanged(object sender, EventArgs e)
        {
            string text = comboBox13.Text;
            sqlCommand = new SqlCommand(application1, sqlConnection);
            try
            {
                sqlReader = sqlCommand.ExecuteReader();
                while (sqlReader.Read())
                {
                    if (Convert.ToString(sqlReader["Application_id"]) == text)
                    {
                        textBox66.Text = Convert.ToString(sqlReader["Guest_name"]);
                        textBox58.Text = Convert.ToString(sqlReader["Room_number"]);
                        textBox57.Text = Convert.ToString(sqlReader["Type"]);
                        textBox56.Text = Convert.ToString(sqlReader["Submission_date"]);
                        textBox55.Text = Convert.ToString(sqlReader["Submission_time"]);
                        comboBox15.Text = Convert.ToString(sqlReader["Execution_status"]);
                        textBox43.Text = Convert.ToString(sqlReader["Guest_id"]);
                        textBox65.Text = Convert.ToString(sqlReader["Completion_date"]);
                        textBox44.Text = Convert.ToString(sqlReader["Completion_time"]);
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

        private void button18_Click(object sender, EventArgs e) //кнопка редактирвоания заявки
        {            
            string text = comboBox13.Text;
            bool app = Methods.Check(application1, text, "Application_id");
            if (app == true) 
            {
                if (!string.IsNullOrEmpty(comboBox13.Text) && !string.IsNullOrEmpty(textBox66.Text) && !string.IsNullOrEmpty(textBox58.Text)
                    && !string.IsNullOrEmpty(textBox57.Text) && !string.IsNullOrEmpty(textBox56.Text) && !string.IsNullOrEmpty(textBox55.Text)
                    && !string.IsNullOrEmpty(comboBox15.Text) && !string.IsNullOrEmpty(textBox43.Text))
                {
                    string name = textBox66.Text;
                    bool guest_name = Methods.Check(guest1, name, "Name");
                    if (guest_name == true)
                    {
                        string room_numb = textBox58.Text;
                        bool room_num = Methods.Check(room1, room_numb, "Room_number");
                        if (room_num == true)
                        {
                            string app_type = textBox57.Text;
                            bool app_type1 = Methods.Check(service1, app_type, "Service_name");
                            if (app_type1 == true)
                            {
                                if (comboBox15.Text == "ожидание принятия" || comboBox15.Text == "принято к выполнению" || comboBox15.Text == "выполнено")
                                {
                                    if (comboBox15.Text == "ожидание принятия" || comboBox15.Text == "принято к выполнению")
                                    {
                                        sqlCommand = new SqlCommand(@"UPDATE Application 
                                                        SET Guest_name = @Guest_name,
                                                        Room_number = @Room_number,
                                                        Type = @Type,
                                                        Execution_status = @Execution_status
                                                        WHERE Application_id = @Application_id", sqlConnection);
                                        sqlCommand.Parameters.AddWithValue("@Guest_name", Convert.ToString(textBox66.Text));
                                        sqlCommand.Parameters.AddWithValue("@Room_number", Convert.ToInt32(textBox58.Text));
                                        sqlCommand.Parameters.AddWithValue("@Type", Convert.ToString(textBox57.Text));
                                        sqlCommand.Parameters.AddWithValue("@Execution_status", Convert.ToString(comboBox15.Text));
                                        sqlCommand.Parameters.AddWithValue("@Application_id", Convert.ToInt32(comboBox13.Text));
                                        sqlCommand.ExecuteNonQuery();
                                        MessageBox.Show("Редактирование прошло успешно!");
                                        Methods.Reload(application, dataGridView5);
                                    }

                                    else if (comboBox15.Text == "выполнено")
                                    {
                                        if (!string.IsNullOrEmpty(textBox65.Text) && !string.IsNullOrEmpty(textBox44.Text))
                                        {
                                            if (Methods.DateChek(textBox65.Text) == true)
                                            {
                                                if (Methods.TimeCheck(textBox44.Text) == true)
                                                {
                                                    sqlCommand = new SqlCommand(@"UPDATE Application 
                                                                SET Guest_name = @Guest_name,
                                                                Room_number = @Room_number,
                                                                Type = @Type,
                                                                Execution_status = @Execution_status,
                                                                Completion_date = @Completion_date,
                                                                Completion_time = @Completion_time
                                                                WHERE Application_id = @Application_id", sqlConnection);
                                                    sqlCommand.Parameters.AddWithValue("@Guest_name", Convert.ToString(textBox66.Text));
                                                    sqlCommand.Parameters.AddWithValue("@Room_number", Convert.ToInt32(textBox58.Text));
                                                    sqlCommand.Parameters.AddWithValue("@Type", Convert.ToString(textBox57.Text));
                                                    sqlCommand.Parameters.AddWithValue("@Execution_status", Convert.ToString(comboBox15.Text));
                                                    sqlCommand.Parameters.AddWithValue("@Completion_date", DateTime.Parse(textBox65.Text));
                                                    sqlCommand.Parameters.AddWithValue("@Completion_time", DateTime.Parse(textBox44.Text));
                                                    sqlCommand.Parameters.AddWithValue("@Application_id", Convert.ToInt32(comboBox13.Text));
                                                    sqlCommand.ExecuteNonQuery();
                                                    MessageBox.Show("Редактирование прошло успешно!");
                                                    Methods.Reload(application, dataGridView5);
                                                }
                                                else
                                                    MessageBox.Show("Время выполнения заявки имеет неверный формат. Введите дату в формате 'дд.мм.гггг'");
                                            }
                                            else
                                                MessageBox.Show("Дата выполнения заявки имеет неверный формат. Введите дату в формате 'дд.мм.гггг'");
                                        }
                                        else
                                            MessageBox.Show("При выбранном статусе 'выполнено' поля 'Дата выполнения' и 'Время выполнения' должны быть заполнены!");
                                    }
                                    else
                                    MessageBox.Show("Выберите статус заявки!");
                                    
                                }
                                else
                                    MessageBox.Show("Выбранного статуса заявки нет в базе данных. Проверьте введенные данные и попробуйте снова.");
                            }
                            else
                                MessageBox.Show("Выбранного типа заявки нет в базе данных. Проверьте введенные данные и попробуйте снова.");
                        }
                        else
                            MessageBox.Show("Выбранной комнаты нет в базе данных. Проверьте введенные данные и попробуйте снова.");
                    }
                    else
                        MessageBox.Show("Такого гостя нет в базе данных. Проверьте введенные данные и попробуйте снова.");
                }
                else
                    MessageBox.Show("Поля должны быть заполнены!");
            }
            else
                MessageBox.Show("Такой заявки нет, проверьте поле 'Номер'!");
        }

        private void comboBox15_SelectedIndexChanged(object sender, EventArgs e) //нужно для редактирования
        {
            if (comboBox15.Text == "ожидание принятия" || comboBox15.Text == "принято к выполнению")
            {
                textBox65.ReadOnly = true;
                textBox44.ReadOnly = true;
            }
            else
            {
                textBox65.ReadOnly = false;
                textBox44.ReadOnly = false;
            }
        }

        private void button19_Click(object sender, EventArgs e) //добавление в таблицу 
        {
            if (!string.IsNullOrEmpty(comboBox14.Text) && !string.IsNullOrEmpty(textBox59.Text) && !string.IsNullOrEmpty(textBox68.Text)
                    && !string.IsNullOrEmpty(comboBox16.Text) && !string.IsNullOrEmpty(textBox64.Text) && !string.IsNullOrEmpty(textBox63.Text)
                    && !string.IsNullOrEmpty(comboBox12.Text))
            {
                string app_type = comboBox16.Text;
                bool app_type1 = Methods.Check(service1, app_type, "Service_name");
                if (app_type1 == true)
                {
                    if (Methods.DateChek(textBox64.Text) == true)
                    {
                        if (Methods.TimeCheck(textBox63.Text) == true)
                        {
                            if (comboBox12.Text == "ожидание принятия" || comboBox12.Text == "принято к выполнению" || comboBox12.Text == "выполнено")
                            {
                                try
                                {
                                    sqlCommand = new SqlCommand(@"INSERT INTO Application (Guest_name, Room_number, Type, Submission_date, Submission_time, Execution_status, Guest_id)
                                                            Values (@Guest_name, @Room_number, @Type, @Submission_date, @Submission_time, @Execution_status, @Guest_id)", sqlConnection);
                                    sqlCommand.Parameters.AddWithValue("@Guest_name", Convert.ToString(textBox59.Text));
                                    sqlCommand.Parameters.AddWithValue("@Room_number", Convert.ToInt32(textBox68.Text));
                                    sqlCommand.Parameters.AddWithValue("@Type", Convert.ToString(comboBox16.Text));
                                    sqlCommand.Parameters.AddWithValue("@Submission_date", DateTime.Parse(textBox64.Text));
                                    sqlCommand.Parameters.AddWithValue("@Submission_time", DateTime.Parse(textBox63.Text));
                                    sqlCommand.Parameters.AddWithValue("@Execution_status", Convert.ToString(comboBox15.Text));
                                    sqlCommand.Parameters.AddWithValue("@Guest_id", Convert.ToInt32(comboBox14.Text));
                                    sqlCommand.ExecuteNonQuery();
                                    MessageBox.Show("Заявка успешно добавлена!");
                                    this.applicationTableAdapter.Fill(this.myHotelDataSet5.Application);
                                    Methods.Reload(application, dataGridView5);
                                }
                                catch { MessageBox.Show("Такого гостя не существует!"); }
                            }
                            else
                                MessageBox.Show("Выбранного статуса заявки нет в базе данных. Проверьте введенные данные и попробуйте снова.");
                        }
                        else
                            MessageBox.Show("Время выполнения заявки имеет неверный формат. Введите дату в формате 'дд.мм.гггг'");
                    }
                    else
                        MessageBox.Show("Дата выполнения заявки имеет неверный формат. Введите дату в формате 'дд.мм.гггг'");
                }
                else
                    MessageBox.Show("Выбранного типа заявки нет в базе данных. Проверьте введенные данные и попробуйте снова.");
            }
            else
                MessageBox.Show("Поля должны быть заполнены!");
        }

        private void comboBox14_SelectedIndexChanged(object sender, EventArgs e) //комбобокс для добавления 
        {
            string text = comboBox14.Text;
            sqlCommand = new SqlCommand(@"Select * from Guest", sqlConnection);
            try
            {
                sqlReader = sqlCommand.ExecuteReader();
                while (sqlReader.Read())
                {
                    if (Convert.ToString(sqlReader["Guest_id"]) == text)
                    {
                        textBox59.Text = Convert.ToString(sqlReader["Name"]);
                        textBox68.Text = Convert.ToString(sqlReader["Room_number"]);
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
           
        private void radioButton18_CheckedChanged(object sender, EventArgs e)
        {
            panel17.Visible = false;
            panel18.Visible = false;
            panel19.Visible = true;
        }

        private void radioButton17_CheckedChanged(object sender, EventArgs e)
        {
            panel17.Visible = false;
            panel19.Visible = false;
            panel18.Visible = true;
        }

        private void radioButton16_CheckedChanged(object sender, EventArgs e)
        {
            panel18.Visible = false;
            panel19.Visible = false;
            panel17.Visible = true;
        }

        private void button30_Click(object sender, EventArgs e)
        {
            Methods.Reload(service, dataGridView6);
        }

        private void button27_Click(object sender, EventArgs e)
        {
            int id = Convert.ToInt32(comboBox17.Text);
            Methods.Delete(delete_service, id);
            Methods.Reload(service, dataGridView6);
            this.serviceTableAdapter1.Fill(this.myHotelDataSet7.Service);
        }
        
        private void comboBox19_SelectedIndexChanged(object sender, EventArgs e) //комбобокс услуги
        {
            string text = comboBox19.Text;
            sqlCommand = new SqlCommand(service1, sqlConnection);
            try
            {
                sqlReader = sqlCommand.ExecuteReader();
                while (sqlReader.Read())
                {
                    if (Convert.ToString(sqlReader["Service_id"]) == text)
                    {
                        textBox60.Text = Convert.ToString(sqlReader["Service_name"]);
                        richTextBox1.Text = Convert.ToString(sqlReader["Description"]);
                        textBox61.Text = Convert.ToString(sqlReader["Price"]);
                        textBox67.Text = Convert.ToString(sqlReader["Realisation_time"]);
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

        private void button28_Click(object sender, EventArgs e) //кнопка редакирования услуги
        {
            if (!string.IsNullOrEmpty(comboBox19.Text) && !string.IsNullOrEmpty(textBox60.Text) && !string.IsNullOrEmpty(richTextBox1.Text)
                    && !string.IsNullOrEmpty(textBox61.Text) && !string.IsNullOrEmpty(textBox67.Text))
            {
                string name = textBox60.Text;
                bool serv_name = Methods.Check(service1, name, "Service_name");
                if (serv_name == true)
                {
                    if (double.TryParse(textBox61.Text, out double Price))
                    {
                        sqlCommand = new SqlCommand(@"UPDATE Service 
                                                        SET Service_name = @Service_name,
                                                        Description = @Description,
                                                        Price = @Price,
                                                        Realisation_time = @Realisation_time
                                                        WHERE Service_id = @Service_id", sqlConnection);
                        sqlCommand.Parameters.AddWithValue("@Service_name", Convert.ToString(textBox60.Text));
                        sqlCommand.Parameters.AddWithValue("@Description", Convert.ToString(richTextBox1.Text));
                        sqlCommand.Parameters.AddWithValue("@Price", Convert.ToDouble(textBox61.Text));
                        sqlCommand.Parameters.AddWithValue("@Realisation_time", Convert.ToString(textBox67.Text));
                        sqlCommand.Parameters.AddWithValue("@Service_id", Convert.ToInt32(comboBox19.Text));
                        sqlCommand.ExecuteNonQuery();
                        MessageBox.Show("Редактирование прошло успешно!");
                        Methods.Reload(service, dataGridView6);
                    }
                    else
                        MessageBox.Show("Поле 'Цена' имеет неверный формат значений!");
                }
                else
                    MessageBox.Show("Такой услуги в базе данных нет!");
            }
            else
                MessageBox.Show("Поля должны быть заполнены!");
        }

        private void button29_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox69.Text) && !string.IsNullOrEmpty(richTextBox2.Text)
                   && !string.IsNullOrEmpty(textBox62.Text) && !string.IsNullOrEmpty(textBox70.Text))
            {
                if (double.TryParse(textBox62.Text, out double Price))
                {
                    sqlCommand = new SqlCommand(@"insert into Service (Service_name, Description, Price, Realisation_time)
                                             VALUES(@Service_name, @Description, @Price, @Realisation_time)", sqlConnection);
                    sqlCommand.Parameters.AddWithValue("@Service_name", Convert.ToString(textBox69.Text));
                    sqlCommand.Parameters.AddWithValue("@Description", Convert.ToString(richTextBox2.Text));
                    sqlCommand.Parameters.AddWithValue("@Price", Convert.ToDouble(textBox62.Text));
                    sqlCommand.Parameters.AddWithValue("@Realisation_time", Convert.ToString(textBox70.Text));
                    sqlCommand.ExecuteNonQuery();
                    MessageBox.Show("Услуга успешно добавлена!");
                    this.serviceTableAdapter1.Fill(this.myHotelDataSet7.Service);
                    Methods.Reload(service, dataGridView6);
                }
                else
                    MessageBox.Show("Поле 'Цена' имеет неверный формат значений!");            
            }
            else
                MessageBox.Show("Поля должны быть заполнены!");
        }

        private void button35_Click(object sender, EventArgs e) //обновить empapp
        {
            Methods.Reload(empapp, dataGridView7);
        }

        private void button40_Click(object sender, EventArgs e) //обновить roomserv
        {
            Methods.Reload(roomserv, dataGridView8);
        }

        private void button21_Click(object sender, EventArgs e) //эксель для организацию
        {
            Methods.SaveExcel(dataGridView1);
        }

        private void button22_Click(object sender, EventArgs e) // эксель для Сотрудников
        {
            Methods.SaveExcel(dataGridView2);
        }

        private void button23_Click(object sender, EventArgs e) //эксель для комнат
        {
            Methods.SaveExcel(dataGridView3);
        }

        private void button24_Click(object sender, EventArgs e) //эксель для посетителей
        {
            Methods.SaveExcel(dataGridView4);
        }

        private void button25_Click(object sender, EventArgs e) // эксель для заявок
        {
            Methods.SaveExcel(dataGridView5);
        }
        private void button26_Click(object sender, EventArgs e) //эксель дл заявок
        {
            Methods.SaveExcel(dataGridView6);
        }

        private void button31_Click(object sender, EventArgs e) //эксель empapp
        {
            Methods.SaveExcel(dataGridView7);
        }

        private void button36_Click(object sender, EventArgs e) //эксель roomserv
        {
            Methods.SaveExcel(dataGridView8);
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Form1 form = new Form1();
            form.Visible = true;
            this.Close();
        }

        private void button33_Click(object sender, EventArgs e)
        {
            try
            {
                int id = Convert.ToInt32(textBox71.Text);
                string query = $@"Select 
                    o.Org_name as 'Организация',
                    o.Organisation_id as 'Номер отдела',
                    o.Adress as 'Адрес отдела',
                    e.Full_name as 'ФИО сотрудника',
                    e.Phone_number as 'Номер телефона',
                    e.Post as 'Должность'
                    from Organization as o 
                    join Employee as e on o.Organisation_id = e.Organisation_id
                    where o.Organisation_id = '{id}'";
                Methods.Reload(query, dataGridView9);
            }
            catch
            {
                MessageBox.Show("Проверьте введенные значения и повторите попытку!");
            }
        }

        private void button32_Click(object sender, EventArgs e)
        {
            Methods.SaveExcel(dataGridView9);
        }

        private void button34_Click(object sender, EventArgs e)
        {
            try
            {
                int id = Convert.ToInt32(textBox72.Text);
                string query = $@"Select 
                o.Org_name as 'Организация',
                o.Organisation_id as 'Номер отдела',
                o.Adress as 'Адрес отдела',
                r.Room_number as 'Номер комнаты',
                r.Room_class as 'Класс комнаты',
                r.Amount_sleeping_place as 'Количество спальных мест',
                g.Name as 'Персона'
                FROM MyHotel.dbo.Room as r 
                join Organization as o on r.Organisation_id = o.Organisation_id
                join Guest as g on r.Room_number = g.Room_number
                WHERE r.Room_number = '{id}'";
                Methods.Reload(query, dataGridView9);
            }
            catch
            {
                MessageBox.Show("Проверьте введенные значения и повторите попытку!");
            }
        }

        private void button37_Click(object sender, EventArgs e)
        {
            try
            {
                int id = Convert.ToInt32(textBox73.Text);
                string query = $@"Select
                    e.Full_name as 'Имя',
                    e.Post as 'Должность',
                    a.Application_id as 'Номер заявки',
                    a.Type as 'Тип заявки',
                    a.Room_number as 'Номер комнаты',
                    a.Guest_name as 'Имя гостя'
                    from EmpApp as ea 
                    join Employee as e on ea.Employee_id = e.Employee_id
                    join Application as a on ea.Application_id = a.Application_id
                    where
                    a.Application_id = '{id}'";
                Methods.Reload(query, dataGridView9);
            }
            catch
            {
                MessageBox.Show("Проверьте введенные значения и повторите попытку!");
            }
        }

        private void button41_Click(object sender, EventArgs e)
        {
            try
            {
                int id = Convert.ToInt32(textBox76.Text);
                string status = Convert.ToString(textBox78.Text);
                string type = Convert.ToString(textBox77.Text);
                string query = $@"Select 
                    o.Org_name as 'Организация',
                    o.Organisation_id as 'Номер отдела',
                    o.Adress as 'Адрес отдела',
                    a.Type as 'Тип заявки',
                    COUNT(ea.Execution_status) as 'Выполнено',
                        (
                        SELECT STRING_AGG(Room_number, ', ')
                        FROM Organization as o
                        JOIN Employee as e ON O.Organisation_id = E.Organisation_id
                        JOIN EmpApp as ea ON EA.Employee_id = E.Employee_id
                        JOIN Application as a ON A.Application_id = EA.Application_id
                        WHERE
                        O.Organisation_id = '{id}'
                        AND EA.Execution_status = '{status}'
                        AND A.Type = '{type}' 
                        )  AS 'Комнаты'
                    from Organization as o
                    JOIN Employee as e ON O.Organisation_id = E.Organisation_id
                    JOIN EmpApp as ea ON EA.Employee_id = E.Employee_id
                    JOIN Application as a ON A.Application_id = EA.Application_id
                    WHERE
                    O.Organisation_id = '{id}'
                    AND EA.Execution_status = '{status}'
                    AND A.Type = '{type}'
                    group by o.Organisation_id, o.Org_name, o.Adress, a.Type";
                Methods.Reload(query, dataGridView9);
            }
            catch
            {
                MessageBox.Show("Проверьте введенные значения и повторите попытку!");
            }
        }

        private void button38_Click(object sender, EventArgs e)
        {
            try
            {
                int id = Convert.ToInt32(textBox74.Text);
                string query = $@"Select 
                    o.Adress as 'Адрес отдела',
                    g.[Name] as 'ФИО',
                    g.Room_number as 'Номер комнаты',
                    (
                     SELECT 
                        CASE 
                            WHEN COUNT(a.Guest_id) = 0 
                                THEN 0 
                                ELSE COUNT(a.Guest_id) END 
                     FROM Organization as o
                     join Room as r on o.Organisation_id = r.Organisation_id
                     join Guest as g on r.Room_number = g.Room_number
                     join Application as a on a.Guest_id = g.Guest_id 
                     Where g.Guest_id = '{id}' 
                    ) AS 'Количество оставленных заявок',
                    g.Settlement_date as 'Дата заселения',
                    g.Settlement_time as 'Время заселения',
                    g.Eviction_date as 'Дата выселения',
                    g.Eviction_time as 'Время выселения'
                    From Organization as o
                    join Room as r on o.Organisation_id = r.Organisation_id
                    join Guest as g on r.Room_number = g.Room_number
                    Where g.Guest_id = '{id}'
                    group by g.[Name], o.Adress, g.Room_number, g.Settlement_date, g.Settlement_time, g.Eviction_date, g.Eviction_time";
                Methods.Reload(query, dataGridView9);
            }
            catch
            {
                MessageBox.Show("Проверьте введенные значения и повторите попытку!");
            }
        }

        private void button39_Click(object sender, EventArgs e)
        {
            try
            {
                string type = Convert.ToString(textBox75.Text);
                string query = $@"Select 
                    s.Service_name as 'Название услуги',
                    Count(a.Type) as 'Количество исполнений',
                    ea.Employee_name as 'Имя сотрудника',
                    (
                        SELECT STRING_AGG(Room_number, ', ')
                        FROM Organization as o
                        JOIN Employee as e ON O.Organisation_id = E.Organisation_id
                        JOIN EmpApp as ea ON EA.Employee_id = E.Employee_id
                        JOIN Application as a ON A.Application_id = EA.Application_id
                        WHERE A.Type = '{type}' 
                    )  AS 'Комнаты'
                    from Service as s
                    join Application as a on a.[Type] = s.Service_Name
                    join EmpApp as ea on a.Application_id = ea.Application_id
                    where a.Type = '{type}'
                    group by s.Service_name, ea.Employee_name";
                Methods.Reload(query, dataGridView9);
            }
            catch
            {
                MessageBox.Show("Проверьте введенные значения и повторите попытку!");
            }
        }
    }
}
