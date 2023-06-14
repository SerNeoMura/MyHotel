using Excel = Microsoft.Office.Interop.Excel;
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
    public class Methods
    {
        //вывод таблицы в datagridview  и кнопка обновления
        public static void Reload(string query, DataGridView dataGridView)
        {
            SqlConnection sqlConnection = new SqlConnection(@"Data Source=ERVOPC;Initial Catalog=MyHotel;Integrated Security=True");
            sqlConnection.Open();
            SqlCommand zapros = new SqlCommand(query, sqlConnection);
            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(zapros);
            System.Data.DataTable table = new System.Data.DataTable();
            sqlDataAdapter.Fill(table);
            dataGridView.DataSource = table;
            dataGridView.ReadOnly = true;
            DataGridViewCellStyle style = dataGridView.ColumnHeadersDefaultCellStyle;
            style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            DataGridViewCellStyle style1 = dataGridView.DefaultCellStyle;
            style1.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
        }

        //проверка для реализации редактирования, чтобы в комбобокс не было вписано несуществующее значение
        public static bool Check(string query, string combtext, string row)
        {
            SqlConnection sqlConnection = new SqlConnection(@"Data Source=ERVOPC;Initial Catalog=MyHotel;Integrated Security=True");
            sqlConnection.Open();
            SqlDataReader sqlReader = null;
            string text = combtext;
            bool org = false;
            SqlCommand sqlCommand = new SqlCommand(query, sqlConnection);
            try
            {
                sqlReader = sqlCommand.ExecuteReader();
                while (sqlReader.Read())
                {
                    if (Convert.ToString(sqlReader[$"{row}"]) == text)
                    {
                        org = true;
                        return org;
                    }

                    else
                        continue;
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
            return org;
        }

        //операция удаления
        public static void Delete(string delete, int text)
        {
            SqlConnection sqlConnection = new SqlConnection(@"Data Source=ERVOPC;Initial Catalog=MyHotel;Integrated Security=True");
            sqlConnection.Open();
            DialogResult dialogResult = MessageBox.Show("Выполнение операции приведет к удалению\n" +
               "всех зависимых записей без возможности восставновления!\n" +
               "Вы уверены?", "Предупреждение", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                try
                {
                    SqlCommand sqlCommand = new SqlCommand(delete, sqlConnection);
                    sqlCommand.Parameters.AddWithValue("@text", text);
                    sqlCommand.ExecuteNonQuery();
                    MessageBox.Show("Запись успешно адалена!");
                }
                catch (Exception exep)
                {
                    MessageBox.Show(exep.Message);
                }
            }
        }

        public static bool PhoneCheked(string text)
        {
            bool chek = true;
            foreach (char sym in text)
            {
                if (char.IsLetter(sym) == true || char.IsPunctuation(sym) == true || char.IsWhiteSpace(sym) == true)
                {
                    chek = false;
                    return chek;
                }
                else
                    continue;
            }
            return chek;
        }

        public static bool Amount_sleep(int amount, int beg, int child)
        {
            bool good = false;
            int sum = beg + child;
            if (sum == amount)
            {
                good = true;
                return good;
            }
            else
                return good;
        }

        public static bool DateChek(string date)
        {
            bool good = false;
            DateTime dDate;

            if (DateTime.TryParse(date, out dDate))
            {
                String.Format("{0:d.MM.yyyy}", dDate);
                good = true;
                return good;
            }
            else
            {
                return good;
            }

        }


        public static bool TimeCheck(string date)
        {
            bool good = false;
            DateTime dDate;

            if (DateTime.TryParse(date, out dDate))
            {
                String.Format("{0:HH:mm:}", dDate);
                good = true;
                return good;
            }
            else
            {
                return good;
            }
        }

        public static void SaveExcel(DataGridView dataGridView) // Сохранить в Excel
        {
            try
            {
                Excel.Application excelapp = new Excel.Application();

                Excel.Workbook workbook = excelapp.Workbooks.Add();
                Excel.Worksheet worksheet = workbook.ActiveSheet;
                for (int i = 1; i < dataGridView.ColumnCount + 1; i++)
                {
                    worksheet.Rows[1].Columns[i] = dataGridView.Columns[i - 1].HeaderCell.Value;
                }
                for (int i = 2; i < dataGridView.RowCount + 2; i++)
                {
                    for (int j = 1; j < dataGridView.ColumnCount + 1; j++)
                    {
                        worksheet.Rows[i].Columns[j] = dataGridView.Rows[i - 2].Cells[j - 1].Value;
                    }
                }

                excelapp.AlertBeforeOverwriting = false;

                excelapp.Visible = true;
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Stop); }
        }
    }
}
