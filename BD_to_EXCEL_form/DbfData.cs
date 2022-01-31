using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Npgsql;
using DotNetDBF;
using System.Data.OleDb;
using System.IO;

namespace BD_to_EXCEL_form
{
    public partial class DbfData : Form
    {
        public static int ID_period { get; set; }
        public DbfData()
        {
            InitializeComponent();

            comboBox1.Items.Clear();
            NpgsqlCommand cmd = new NpgsqlCommand($"SELECT id, Period_name from regop_period order by 1 desc", Form1.con);
            using (NpgsqlDataReader ndr = cmd.ExecuteReader())
            {
                if (ndr.HasRows) // если есть данные
                {
                    while (ndr.Read()) // построчно считываем данные
                    {
                        comboBox1.Items.Add(ndr.GetValue(1));
                    }
                }
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            NpgsqlCommand cmd = new NpgsqlCommand($"select id from regop_period where Period_name='{comboBox1.SelectedItem.ToString()}'", Form1.con);
            ID_period = (int)Convert.ToInt64(cmd.ExecuteScalar().ToString());
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DataTable selDt = new DataTable();
            NpgsqlDataAdapter sda = new NpgsqlDataAdapter($@"select acc_num ACC_CODE, 
                                        b4fa.street_name as STREET,
                                        b4f.shortname as STR_TYPE,
                                        b4fa.house as HOUSE,
                                        croom_num as FLAT_NUM,
                                        LTRIM(own.name, ' ') as MASTER,
                                        to_char(rpa.open_date, 'dd.mm.yyyy'),
                                        to_char(rpa.close_date, 'dd.mm.yyyy'),
                                        null AS PHONE,
                                        case 
                                        	when gr.ownership_type = 10 then 'Частная' 
                                        	when gr.ownership_type = 30 then 'Муниципальная' 
                                        	when gr.ownership_type = 40 then 'Государственная'
                                        	when gr.ownership_type = 50 then 'Коммерческая' 
                                        	when gr.ownership_type = 60 then 'Смешанная' 
                                        	when gr.ownership_type = 80 then 'Федеральная'
                                        	when gr.ownership_type = 90 then 'Областная' 
                                        	else 'Не указано' 
                                        end as OWNERSHIP,
                                        case 
                                        	when gr.type = 10 and gr.is_communal = true then 'Коммунальное помещение'
                                        	when gr.type = 10 then 'Жилое помещение'
                                        	when gr.type = 20 then 'Нежилое помещение' 
                                        	else 'Не указано' 
                                        end as HABIT_TYPE,
                                        area_mkd as TOTAL_SQ,
                                        area_living_owned as LIVING_SQ,
                                        null AS LODGER_CNT
                                        from regop_pers_acc rpa
                                        join gkh_room gr on gr.id=rpa.room_id
                                        join regop_pers_acc_owner own on rpa.acc_owner_id = own.id and rpa.state_id = 804
                                        join gkh_reality_object ro on ro.id = gr.ro_id
                                        join b4_fias_address b4fa on b4fa.id= ro.fias_address_id
                                        join b4_fias b4f on b4f.aoguid = b4fa.street_guid and b4f.actstatus=1
                                        where municipality_id=4808
                                        order by 2,4,5,6", Form1.con);
            sda.Fill(selDt);

            label2.Text="Количество записей: " + selDt.Rows.Count.ToString();

            for (int i = 0; i < selDt.Rows.Count + 1; i++)
            {
                Console.WriteLine("Начинаю файл " + i);
                OleDbConnection conn = new OleDbConnection();

                conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Application.StartupPath + @"\Data;Extended Properties=dBASE IV;User ID=Admin;";
                conn.Open();
                OleDbCommand comm = conn.CreateCommand();
                comm.CommandText = "DELETE * FROM NEW";
                comm.ExecuteNonQuery();
                conn.Close();
                label2.Text = System.IO.Path.GetDirectoryName(Application.ExecutablePath) + @"\New.dbf";
                FileInfo fn = new FileInfo(System.IO.Path.GetDirectoryName(Application.ExecutablePath) + @"\Data\New.dbf");
                fn.CopyTo(System.IO.Path.GetDirectoryName(Application.ExecutablePath) + @"\New.dbf", true);
                conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Application.StartupPath + ";Extended Properties=dBASE IV;User ID=Admin;";
                conn.Open();
                comm.CommandText = "INSERT INTO NEW VALUES('test','test', 'test','test','test','test','test','test','test','t','1','test','test','test','test','1','1','test','test','1','1','test','1','1','1','1','1','1','1','1','1','test','test','test','test','test','test','test','test')";
                comm.ExecuteNonQuery();
                /*while (selDt.Rows[i][0].ToString() == mainDt.Rows[ji][0].ToString())
                {
                    try
                    {
                        //comm.CommandText = "INSERT INTO SHABLON VALUES('34','Региональный','6806','Октябрьское','с','ул','Ленина','57','','12','2018','8','капитальный ремонт','капитальный ремонт','кв.м',7.400,23.02,123.20,6563.54,6663.25,5305.40,3)";
                        comm.CommandText = "INSERT INTO SHABLON VALUES('" + mainDt.Rows[ji][0].ToString() + "'" + delimeter + "'" + mainDt.Rows[ji][1].ToString() + "'" + delimeter + "'" + mainDt.Rows[ji][2].ToString()
                            + "'" + delimeter + "'" + mainDt.Rows[ji][3].ToString() + "'" + delimeter + "'" + mainDt.Rows[ji][4].ToString() + "'" + delimeter + "'" + mainDt.Rows[ji][5].ToString() + "'"
                            + delimeter + "'" + mainDt.Rows[ji][6].ToString() + "'" + delimeter + "'" + mainDt.Rows[ji][7].ToString() + "'" + delimeter + "'" + mainDt.Rows[ji][8].ToString() + "'"
                            + delimeter + "'" + mainDt.Rows[ji][9].ToString() + "'" + delimeter + "'" + mainDt.Rows[ji][10].ToString() + "'" + delimeter + "'" + mainDt.Rows[ji][11].ToString() + "'"
                            + delimeter + "'" + mainDt.Rows[ji][12].ToString() + "'" + delimeter + "'" + mainDt.Rows[ji][13].ToString() + "'" + delimeter + "'" + mainDt.Rows[ji][14].ToString() + "'"
                            + delimeter + mainDt.Rows[ji][15].ToString() + delimeter
                            + Decimal.Round(Convert.ToDecimal(mainDt.Rows[ji][16].ToString()), 2).ToString().Replace(',', '.') + delimeter
                            + Decimal.Round(Convert.ToDecimal(mainDt.Rows[ji][17].ToString()), 2).ToString().Replace(',', '.') + delimeter
                            + Decimal.Round(Convert.ToDecimal(mainDt.Rows[ji][18].ToString()), 2).ToString().Replace(',', '.') + delimeter
                            + Decimal.Round(Convert.ToDecimal(mainDt.Rows[ji][19].ToString()), 2).ToString().Replace(',', '.') + delimeter
                            + Decimal.Round(Convert.ToDecimal(mainDt.Rows[ji][20].ToString()), 2).ToString().Replace(',', '.') + delimeter
                            + Convert.ToInt32(mainDt.Rows[ji][21].ToString()).ToString() + ")";
                        comm.ExecuteNonQuery();
                        if (ji < mainDt.Rows.Count - 1)
                        {
                            ji++;
                        }
                        else
                        {
                            break;
                        }
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                        lf.WriteLine(e.Message);
                    }
                }*/
                conn.Close();
                //File.Move(@"D:\temp\sobits\1234\SHABLON.dbf", @"D:\temp\sobits\1234\" + selDt.Rows[i][0].ToString() + "18080000.dbf");
            }
        }
    }
}
