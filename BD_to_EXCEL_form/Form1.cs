using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using BD_to_EXCEL_form;
using System.Reflection;
using Npgsql;
using DocumentFormat.OpenXml;
using ClosedXML.Excel;
using DotNetDBF;
using System.Data.OleDb;

namespace BD_to_EXCEL_form
{
    public partial class Form1 : Form
    {
        public static Dictionary<string, string> alias = new Dictionary<string, string>();
        public static NpgsqlCommand cmd = new NpgsqlCommand();
        public static NpgsqlConnection con = new NpgsqlConnection();
        public static NpgsqlDataReader ndr;

        public Form1()
        {
            InitializeComponent();
            //textBox1.Text = "Server=172.153.153.46;Port=5432;Database=gkh_chelyabinsk;User ID=bars;Password=bars;CommandTimeout=2000000;";
            textBox1.Text = "Server=192.168.1.51;Database=gkh_chelyabinsk;UserID=postgres;Password=1234;CommandTimeout=2000000;";
            
            Dict_read();
            Servers_read();
            //для теста потом снести
            textBox3.Text = @"drop table if exists gis_by_ro;
                            create temp table gis_by_ro as
                            (select * from gkh_reality_object limit 10 );
                            select * from gis_by_ro";
        }
        private void button2_Click(object sender, EventArgs e)
        {
            log("Подключение к: " + textBox1.Text);
            string line_con = textBox1.Text;

            try
            {
                con = new NpgsqlConnection(line_con);
                cmd = new NpgsqlCommand("select count(*) from information_schema.tables where table_schema = 'public';", con);
                con.Open();
                log("Подключение успешно, в схеме 'public':  " + cmd.ExecuteScalar().ToString() + " таблиц!");
            }
            catch (Exception ex)
            {
                log("Ошибка!  " + ex.Message + "\n");
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1 = new OpenFileDialog();
        }
        public async void log(string s)
        {
            textBox2.AppendText(s + "\r\n");
            using (StreamWriter sw = new StreamWriter(System.IO.Path.GetDirectoryName(Application.ExecutablePath) + @"\logs.txt", true, System.Text.Encoding.UTF8))
            {
                await sw.WriteLineAsync(DateTime.UtcNow.ToString() + "|" + s);
            }
        }
        /// <summary>
        /// Чтение файла *\aliases.txt, заполнение справочника c псевдонимами
        /// </summary>
        public async void Dict_read()
        {
            string path = System.IO.Path.GetDirectoryName(Application.ExecutablePath) + @"\aliases.txt";
            log(@"Чтение файла *\aliases.txt");

            start:
            if (File.Exists(System.IO.Path.GetDirectoryName(Application.ExecutablePath) + @"\aliases.txt"))
            {
                using (StreamReader sr = new StreamReader(System.IO.Path.GetDirectoryName(Application.ExecutablePath) + @"\aliases.txt", System.Text.Encoding.UTF8))
                {
                    string line;
                    while ((line = await sr.ReadLineAsync()) != null)
                    {
                        alias.Add(line.Split('|')[0].ToString(), line.Split('|')[1].ToString());
                    }
                    await sr.ReadLineAsync();
                }
            }
            else
            {
                MessageBox.Show("Файл не создан");
                using (StreamWriter sw = new StreamWriter(System.IO.Path.GetDirectoryName(Application.ExecutablePath) + @"\aliases.txt", true, System.Text.Encoding.UTF8))
                {
                    await sw.WriteLineAsync("TEST|ТЕСТ");
                }
                goto start;
            }
        }
        private void textBox2_VisibleChanged(object sender, EventArgs e)
        {
        }
        /// <summary>
        /// Считываем файл *\servers.txt, все найденые строки помещаются в ComboBox
        /// </summary>
        public async void Servers_read()
        {
            string path = System.IO.Path.GetDirectoryName(Application.ExecutablePath) + @"\servers.txt";
            log(@"Чтение файла *\servers.txt");
            start:
            if (File.Exists(System.IO.Path.GetDirectoryName(Application.ExecutablePath) + @"\servers.txt"))
            {
                using (StreamReader sr = new StreamReader(System.IO.Path.GetDirectoryName(Application.ExecutablePath) + @"\servers.txt", System.Text.Encoding.UTF8))
                {
                    string line;
                    while ((line = await sr.ReadLineAsync()) != null)
                    {
                        comboBox1.Items.Add(line);
                        //alias.Add(line.Split('|')[0].ToString(), line.Split('|')[1].ToString());
                    }
                    await sr.ReadLineAsync();
                }
            }
            else
            {
                log(@"Файл  *\servers.txt. Создание файла и переоткрытие...");
                using (StreamWriter sw = new StreamWriter(System.IO.Path.GetDirectoryName(Application.ExecutablePath) + @"\servers.txt", true, System.Text.Encoding.UTF8))
                {
                    await sw.WriteLineAsync("Server=192.168.1.51;Database=gkh_chelyabinsk;UserID=postgres;Password=1234;CommandTimeout=2000000;");
                }
                goto start;
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            button4.Enabled = false;
            button5.Enabled = false;
            button6.Enabled = false;

            dataGridView1.Rows.Clear();
            //ndr=null;
            try
            {
                string sql = textBox3.Text;
                cmd = new NpgsqlCommand(sql, con);
                ndr = cmd.ExecuteReader();
                List<string> headerTable = new List<string>(); //получение заголовков
                for (int i = 0; i < ndr.FieldCount; i++)
                {
                    headerTable.Add(ndr.GetName(i).ToString());
                    dataGridView1.Rows.Add();
                    dataGridView1[1, i].Value = headerTable[i].ToString();
                    if (alias.ContainsKey(headerTable[i].ToString()))
                    {
                        dataGridView1[0, i].Value = true;
                        dataGridView1[2, i].Value = alias[headerTable[i].ToString()];
                    }
                }
                ndr.Close();
                button4.Enabled = true;
                button5.Enabled = true;
                button6.Enabled = true;
            }
            catch (Exception ex)
            {
                log(ex.Message);
            }
        }
        
        /// <summary>
        /// проверяем существование дирректории '*/data@' если нет создаем
        /// </summary>
        private void Directory_data()
        {
            if (!File.Exists(System.IO.Path.GetDirectoryName(Application.ExecutablePath) + @"\data"))
            {
                log("Создаем дирректрию: " + (Application.StartupPath) + @"\data");
                DirectoryInfo di = Directory.CreateDirectory((Application.StartupPath) + @"\data");
            }
        }
        /// <summary>
        /// Получение заголовков
        /// </summary>
        /// <returns></returns>
   
        private void read_csv(string sql)
        {
            Directory_data();
            NpgsqlCommand nc = new NpgsqlCommand(sql, con);
            NpgsqlDataReader ndr = nc.ExecuteReader();
            log("Извлечение данных\nИзвлечение заголовков");
            List<string> headerTable = new List<string>(); //получение заголовков
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (dataGridView1[2, i].Value != null)
                {
                    headerTable.Add(dataGridView1[2, i].Value.ToString());
                }
                else
                {
                    headerTable.Add(dataGridView1[1, i].Value.ToString());
                }
            }

            log("Генерация csv файла, запись данных");
            string dataname = date();
            StreamWriter file = new StreamWriter(System.IO.Path.GetDirectoryName(Application.ExecutablePath) + @"\data\" + dataname + ".csv", false, Encoding.UTF8);
            string strok_csv = "";
            try
            {
                strok_csv = "";
                foreach (string s in headerTable)
                {
                    strok_csv += @"""" + s.ToString() + @""";";
                }
                log("Заполняем заголовки");
                file.WriteLine(strok_csv);

                string temp_strok = "";
                if (ndr.HasRows)
                {
                    while (ndr.Read())
                    {
                        strok_csv = "";
                        for (int x = 0; x < ndr.FieldCount; x++)
                        {
                            try
                            {
                                temp_strok = ndr.GetValue(x).ToString().Replace(@"/n", "").Replace(@"/r", "");
                                //strok_csv += @"""" + temp_strok + @""";";
                                strok_csv += temp_strok + @";";
                            }
                            catch (System.InvalidCastException)
                            {
                                //log(ice.Message);
                                //log("Вероятнее всего проблема с преобразованием даты \"infinity\"");
                                temp_strok = "infinity";
                                //strok_csv += @"""" + temp_strok + @""";";
                                strok_csv += temp_strok + @";";
                            }

                        }
                        file.WriteLine(strok_csv);
                    }
                }
                else
                {
                    log("Не обнаружены строки для записи в csv");
                }
                log("Генерация csv файла, запись данных: ГОТОВО");
                //file.Close();
                //ndr.Close();
                //con.Dispose();
            }
            catch (Exception ex)
            {
                log(ex.Message);
            }
            finally
            {
                ndr.Close();
                file.Close();
                //if (con.State == ConnectionState.Open)con.Close();
                // con.Dispose();
            }
        }
        private void button6_Click(object sender, EventArgs e)
        {
            string filePath = (Application.StartupPath) + @"\data\" + date().ToString() + ".xlsx";
            
            Dict_update();
            log("Проверка существования дирректории: " + (Application.StartupPath) + @"\data\");
            Directory_data();
            log("Открытие файла excel");
            
            XLWorkbook xLWorkbook = new XLWorkbook();
            log("Добавляем заголовки");

            List<string> headerTable = new List<string>(); //получение заголовков
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (dataGridView1[2, i].Value != null)
                {
                    headerTable.Add(dataGridView1[2, i].Value.ToString());
                }
                else
                {
                    headerTable.Add(dataGridView1[1, i].Value.ToString());
                }
            }
            //Добавляем лист "Выгрузка", добавляем заголовки
            var excelworksheet = xLWorkbook.Worksheets.Add("Выгрузка");


            NpgsqlCommand nc = new NpgsqlCommand(textBox3.Text, con);
            NpgsqlDataReader ndr = nc.ExecuteReader();

            //string strok_csv = "";
            try
            {
                for (int i = 0; i < headerTable.Count; i++)
                {
                    excelworksheet.Cell(1, i + 1).Style.Fill.BackgroundColor = XLColor.LemonChiffon;
                    excelworksheet.Columns(1, i + 1).AdjustToContents();
                    excelworksheet.Cell(1, i + 1).Value = headerTable[i].ToString();
                }

                string temp_strok = "";
                if (ndr.HasRows)
                {
                    int z = 2;
                    while (ndr.Read())
                    {
                        for (int x = 0; x < ndr.FieldCount; x++)
                        {
                            try
                            {
                                temp_strok = ndr.GetValue(x).ToString().Replace(@"/n", "").Replace(@"/r", "");
                                excelworksheet.Cell(z, x + 1).Value = temp_strok;
                                //strok_csv += @"""" + temp_strok + @""";";
                            }
                            catch (System.InvalidCastException)
                            {
                                //log(ice.Message);
                                //log("Вероятнее всего проблема с преобразованием даты \"infinity\"");
                                excelworksheet.Cell(z, x + 1).Value = "infinity";
                            }

                        }
                        //file.WriteLine(strok_csv);
                        z++;
                    }
                }
                else
                {
                    log("Не обнаружены строки для записи в csv");
                }
                log("Генерация csv файла, запись данных: ГОТОВО");
                //file.Close();
                //ndr.Close();
                //con.Dispose();
            }
            catch (Exception ex)
            {
                log(ex.Message);
            }
            finally
            {
                ndr.Close();
                //if (con.State == ConnectionState.Open)con.Close();
                // con.Dispose();
            }
            //Console.WriteLine("Записей: " + reccount);

            /*for (int i = 2; i <= Convert.ToInt32(reccount); i++)
            {
                Console.WriteLine((i - 1).ToString() + " Строка");
                //if (excelworksheet.Cells[i, 10].Value.ToString().Contains(","))
                //{
                //    // дополнительный метод для преобразования
                //    string fds = id_ls_join(excelworksheet.Cells[i, 8].Value, excelworksheet.Cells[i, 10].Value, i);
                //}
            
                cmd.CommandText = $@"select gr.croom_num, 
                CASE WHEN ro.gis_gkh_guid is null THEN b4fa.house_guid::text
                WHEN ro.gis_gkh_guid is not null THEN ro.gis_gkh_guid
                ELSE 'Не найден' END, 
                string_agg(acc_num, ', ') as acc_num, string_agg(rpa.id::char(50), ', ')
                from regop_pers_acc rpa
                join gkh_room gr on gr.id = rpa.room_id
                join gkh_reality_object ro on ro.id = gr.ro_id
                join b4_fias_address b4fa on b4fa.id = ro.fias_address_id
                where (b4fa.house_guid::text = '{excelworksheet.Row(i).Cell(8).Value}' and gr.croom_num = '{excelworksheet.Row(i).Cell(10).Value}') or (ro.gis_gkh_guid = '{excelworksheet.Row(i).Cell(8).Value}' and gr.croom_num = '{excelworksheet.Row(i).Cell(10).Value}')
                group by 1,2";
                ndr = cmd.ExecuteReader();
            }
            Console.WriteLine("Сохраняем изменения");*/
            xLWorkbook.SaveAs(filePath);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Dict_update();
            read_csv(textBox3.Text);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Dict_update();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox1.Text = comboBox1.SelectedItem.ToString();
        }

        public void Dict_update()
        {
            try
            {
                foreach (DataGridViewRow dr in dataGridView1.Rows)
                {
                    //MessageBox.Show(dr.Cells[1].Value.ToString() + "///" + dr.Cells[2].Value.ToString());
                    if (alias.ContainsKey(dr.Cells[1].Value.ToString()) && dr.Cells[2].Value != null)
                    {
                        //MessageBox.Show("Найдено:" + dr.Cells[1].Value.ToString() + "///" + dr.Cells[2].Value.ToString());
                        log("Обновлены алиасы: " + dr.Cells[1].Value.ToString());
                        alias[dr.Cells[1].Value.ToString()] = dr.Cells[2].Value.ToString();
                    }
                    else if (!alias.ContainsKey(dr.Cells[1].Value.ToString()) && dr.Cells[2].Value != null)
                    {
                        log("Добавлены алиасы: " + dr.Cells[1].Value.ToString());
                        alias.Add(dr.Cells[1].Value.ToString(), dr.Cells[2].Value.ToString());
                    }
                    else
                    {
                        log("Пропущенны алиасы: " + dr.Cells[1].Value.ToString());
                    }
                }

                ShowIterator(alias);
            }
            catch (Exception ex)
            {
                log(ex.Message);
            }
        }
        public async void ShowIterator<K, V>(Dictionary<K, V> myList)
        {
            string path = System.IO.Path.GetDirectoryName(Application.ExecutablePath) + @"\aliases.txt";
            log(@"Обновление файла *\aliases.txt");

            if (myList == null)
            {
                log("Словарь алиасов вернул ноль значений");
                return;
            }

            if (File.Exists(System.IO.Path.GetDirectoryName(Application.ExecutablePath) + @"\aliases.txt"))
            {
                using (StreamWriter sw = new StreamWriter(System.IO.Path.GetDirectoryName(Application.ExecutablePath) + @"\aliases.txt", false, System.Text.Encoding.UTF8))
                {
                    foreach (KeyValuePair<K, V> kvp in myList)
                    {
                        await sw.WriteLineAsync(kvp.Key.ToString() + "|" + kvp.Value.ToString());
                    }
                }
            }
            else
            {
                MessageBox.Show("Файл не создан");

            }
        }

        private void button7_Click_1(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("explorer", System.IO.Path.GetDirectoryName(Application.ExecutablePath) + @"\data");
        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                cmd = null;
                ndr = null;
                if (con.State == ConnectionState.Open) con.Close();
                con.Dispose();
                log("Соединение и адаптер обнулены и сброшены");
            }
            catch (Exception ex)
            {
                log("При сбросе произошла ошибка: " + ex.Message);
            }

        }

        public string date ()
        {
            DateTime thisDay = DateTime.Today;
            return thisDay.ToString("dd-MM-yyyy");
        }

        private void dBFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (con.State == ConnectionState.Open)
            {
                DbfData df = new DbfData();
                df.ShowDialog();
            }
            else
            {
                log("Для выгрузки DBF нужно активное подключение");
            }
        }
    }
}
