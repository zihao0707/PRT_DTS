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
            var prtData = Get_prtData();
           
        }

        /// <summary>
        /// 取得一筆列印資料
        /// </summary>
        /// <param MACAddress.Text>配對碼</param>
        /// <returns></returns>
        private string Get_prtData()
        {
            string sSQL = @$"select top 1 * FROM PRT01_0000 WHERE  usr_code = {MACAddress.Text}";//取得列印資料

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

            return "列印完畢";
        }
    }
}
