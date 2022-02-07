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
        public static string name_file { get; set; }
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
            cmd.Cancel(); cmd.Dispose();
            cmd = new NpgsqlCommand($"select to_char(cstart, 'MMYY') from regop_period where Period_name='{comboBox1.SelectedItem.ToString()}'", Form1.con);
            name_file = "NACH" + cmd.ExecuteScalar().ToString()+".dbf";
            button1.Enabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            const string quote = "\"";
            
            DataTable selDt = new DataTable();
            NpgsqlDataAdapter sda = new NpgsqlDataAdapter($@"select 
                        unified_acc_num {quote}UNIFIED_AC{quote},rpa.acc_num {quote}ACC_NUMBER{quote}, own.name {quote}OWNER{quote}, gdm.name {quote}MU{quote},
                        replace(split_part(b4fa.place_name, ' ', 1), '.', '') {quote}KINDSITY{quote},
                        SUBSTRING(b4fa.place_name, 4, 999) {quote}CITY{quote},
                        replace(split_part(b4fa.street_name, ' ', 1), '.', '') {quote}KINDSTREET{quote},
                        b4fa.street_name {quote}STREET{quote},
                        b4fa.HOUSE {quote}HOUSE{quote},
                        letter {quote}LETTER{quote},
                        (case when b4fa.housing is not null or b4fa.housing <> '' then null else b4fa.housing end)::numeric(4) as {quote}HOUSING{quote},
                        b4fa.BUILDING {quote}BUILDING{quote}, gr.croom_num {quote}CROOM_NUM{quote}, b4fa.house_guid {quote}ADR_FIAS{quote},'' as {quote}PRIM{quote},
                        round(gr.carea, 2) {quote}PLOSHAD{quote},round(area_living_owned, 2) as {quote}LIVING_SQ{quote},
                        case 
                        when gr.ownership_type = 10 then 'Частная'
                        when gr.ownership_type = 30 then 'Муниципальная'
                        when gr.ownership_type = 40 then 'Государственная'
                        when gr.ownership_type = 50 then 'Коммерческая'
                        when gr.ownership_type = 60 then 'Смешанная'
                        when gr.ownership_type = 80 then 'Федеральная'
                        when gr.ownership_type = 90 then 'Областная'
                        else 'Не указано'
                        end as {quote}OWNERSHIP{quote}, 
                        case when gr.IS_COMMUNAL then 'КОММУНАЛЬНАЯ' else 'ОТДЕЛЬНАЯ' end as {quote}HABIT_TYPE{quote},
                        '' as {quote}PROPIS{quote},'041' {quote}SRV_ID{quote},'' as {quote}SRV_NAME{quote},
                        (1) as {quote}REC_TYPE{quote},
                        '10,10000' {quote}TARIF{quote},
                        '0' {quote}NORM{quote},
                        round(psum.charge_tariff, 2) as {quote}SUMMA{quote},
                        round(psum.RECALC, 2) as {quote}RECALC{quote},
                        round(BASE_TARIFF_DEBT, 2) as {quote}DOLG{quote},
                        round(PENALTY, 2) {quote}PENI{quote},
                        round((PENALTY_PAYMENT + TARIFF_PAYMENT), 2) {quote}OPLATA{quote},
                        round(SALDO_OUT_SERV, 2) {quote}SUMMA_K_OP{quote},
                        '40603810209280004926' as {quote}RS{quote},'Филиал {quote}Центральный{quote} Банка ВТБ (ПАО) в г. Москве' as {quote}BANK{quote},'30101810145250000411' as {quote}KOR{quote},'044525411' as {quote}BIK{quote},'454048' as {quote}INDEX{quote},
                        to_char(rp.cstart, 'MMYY') as {quote}PERIOD_OPL{quote}, To_char(rp.cstart + interval '1 month', 'DDMMYYYY') as {quote}OPLATIT_DO{quote},'fondkp174@mail.ru' as {quote}EMAIL{quote}
                        from regop_pers_acc_period_summ psum
                        join regop_period rp on rp.id = psum.period_id
                        join regop_pers_acc rpa on rpa.id = psum.account_id and rpa.state_id = 804
                        join gkh_room gr on gr.id = rpa.room_id
                        join regop_pers_acc_owner own on rpa.acc_owner_id = own.id and rpa.state_id = 804
                        join gkh_reality_object ro on ro.id = gr.ro_id
                        join gkh_dict_municipality gdm on ro.municipality_id = gdm.id
                        join b4_fias_address b4fa on b4fa.id = ro.fias_address_id
                        where psum.period_id = {ID_period} order by 3 limit 10", Form1.con);
            sda.Fill(selDt);
            label2.Text="Количество записей: " + selDt.Rows.Count.ToString();
            string insert_str = "";

            //selDt.Rows[].Columns[1].,5]
            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Application.StartupPath + @"\Data;Extended Properties=dBASE IV;User ID=Admin;";
            conn.Open();
            OleDbCommand comm = conn.CreateCommand();
            comm.CommandText = "DELETE * FROM NEW";
            comm.ExecuteNonQuery();
            conn.Close();
            //label2.Text = System.IO.Path.GetDirectoryName(Application.ExecutablePath) + @"\New.dbf";
            FileInfo fn = new FileInfo(System.IO.Path.GetDirectoryName(Application.ExecutablePath) + @"\Data\New.dbf");
            fn.CopyTo(System.IO.Path.GetDirectoryName(Application.ExecutablePath) + @"\New.dbf", true);
            conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Application.StartupPath + ";Extended Properties=dBASE IV;User ID=Admin;";
            //comm.CommandText = "INSERT INTO NEW VALUES('test','test', 'test','test','test','test','test','test','test','t','1','test','test','test','test','1','1','test','test','1','1','test','1','1','1','1','1','1','1','1','1','test','test','test','test','test','test','test','test')";
            conn.Open();
            comm = conn.CreateCommand();
            for (int i = 0; i < selDt.Rows.Count; i++)
            {
                insert_str = "";
                for (int j = 0; j < selDt.Columns.Count; j++)
                {
                    insert_str += $"{(selDt.Rows[i][j].ToString()==""?"null":"'"+selDt.Rows[i][j].ToString().Trim()+"'")},";
                }
                comm.CommandText = $"INSERT INTO NEW VALUES({insert_str.Remove(insert_str.Length-1)})";
                comm.ExecuteNonQuery();
                
                //File.Move(@"D:\temp\sobits\1234\SHABLON.dbf", @"D:\temp\sobits\1234\" + selDt.Rows[i][0].ToString() + "18080000.dbf");
            }
            conn.Close();
            MessageBox.Show("End:  "+System.IO.Path.GetDirectoryName(Application.ExecutablePath) + @"\" + name_file);
        }
    }
}
