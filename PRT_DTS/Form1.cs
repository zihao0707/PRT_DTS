using PRT_DTS.PrintData;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PRT_DTS
{
    public partial class Form1 : Form
    {
        Comm _comm = new Comm();
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Comm Comm = new Comm();
            var MAC = Comm.GetMACAddress();
            MACAddress.Text = MAC.Count == 0 ? "" : MAC[0];
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!new TSCLIB().Test_Print()) { label3.Text = ""; };
           
           
        }

        /// <summary>
        /// 取得一筆列印資料
        /// </summary>
        /// <param MACAddress.Text>配對碼</param>
        /// <returns></returns>
        private string Get_prtData()
        {
            string sSQL = @$"select top 1 * FROM PRT01_0000 
                                WHERE  usr_code = '{MACAddress.Text}'";
                                   //and print_name ='{printName.Text}';//取得列印資料

            DataTable prtData = new Comm().Get_DataTable(sSQL, "MES");
            if (prtData.Rows.Count == 0) { return "無資料"; }

            PRT01_0000 prt = new PRT01_0000();
            foreach (DataRow db in prtData.Rows)
            {
                prt.Prt0100001 = _comm.String_ParseInt32(db["prt01_0000"].ToString());
                prt.PrtType = db["prt_type"].ToString();     //列印類別
                prt.PrtKind = db["prt_kind"].ToString();     //列印種類
                prt.PrintName = db["print_name"].ToString(); //印表機名稱
                prt.PrintData = db["print_data"].ToString(); //列印內容
                prt.UsrCode = db["usr_code"].ToString();     //配對碼
                prt.LabelCode = db["label_code"].ToString(); //標籤代碼
            }

            new TSCLIB().Print_LabelsData(prt);
            DataTable insData = new DataTable();
            new Comm().ins_PRT02_0000(prt,"MES");
            Del_prtData(prt);
            return "列印完畢";
        }

        /// <summary>
        /// 取得一筆列印資料
        /// </summary>
        /// <param MACAddress.Text>配對碼</param>
        /// <returns></returns>
        private bool Del_prtData(PRT01_0000 pRT)
        {
            new Comm().Del_QueryData("PRT01_0000", "prt01_0000", new Comm().ParseString(pRT.Prt0100001),"MES");
            return true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            new TSCLIB()._error = "";
            timer1.Enabled = true;
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            
            string err = new TSCLIB()._error.ToString().Length>1
                ? new TSCLIB()._error
                : "";
            if (err.Length>1) {
                MessageBox.Show(err);
                timer1.Enabled = false;
            }
            
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            string sSql = @"delete PRT01_0000";
            new Comm().SaveSQL(sSql, "MES");
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            Get_prtData();
        }
    }
}
