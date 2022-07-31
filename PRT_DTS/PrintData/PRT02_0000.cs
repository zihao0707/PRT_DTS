using System;
using System.Collections.Generic;
using System.Text;

namespace PRT_DTS.PrintData
{
    public class PRT02_0000
    {
        public string prt02_0000 { get; set; }
        public string prt_type { get; set; }
        public string prt_kind { get; set; }
        public string print_name { get; set; }
        public string print_data { get; set; }
        public string result { get; set; }
        public string ins_date { get; set; }
        public string ins_time { get; set; }
        public string usr_code { get; set; }
        public string prt_date { get; set; }
        public string prt_time { get; set; }


        public PRT02_0000 insPRT02_0000(PRT01_0000 prt)
        {
            PRT02_0000 prt02 = new PRT02_0000();
            prt02.prt02_0000 = new Comm().Get_NewGUID();
            prt02.prt_type = prt.PrtType;
            prt02.prt_kind = prt.PrtKind;
            prt02.print_name = prt.PrintName;
            prt02.print_data = prt.PrintData;
            prt02.usr_code = prt.UsrCode;
            prt02.result = "OK";
            prt02.ins_date = DateTime.Now.ToString("yyyy-MM-dd");
            prt02.ins_time = DateTime.Now.ToString("HH:mm:ss");
            prt02.prt_date = DateTime.Now.ToString("yyyy-MM-dd");
            prt02.prt_time = DateTime.Now.ToString("HH:mm:ss");

            return prt02;
        }

    }
}