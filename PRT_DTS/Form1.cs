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
            string sSQL = @$"select * FROM PRT01_0000 WHERE  usrCode = {MACAddress.Text}";//取得列印資料
            DataTable insData = new Comm().Get_DataTable(sSQL, "ERP");

            var result = from p in insData.AsEnumerable()
                         select new
                         {
                             print_data = p.Field<string>("print_data")
                         };
            string print_data = result.ToArray()[0].print_data;
            new TSCLIB().Print_LabelsData(print_data, insData);
        }
    }
}
