using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace PRT_DTS.PrintData
{
    public class OME_QC
    {
        public string TZ { get; set; }
        public string TW { get; set; }
        public string DZ { get; set; }
        public string DW { get; set; }
        public string W1 { get; set; }
        public string W2 { get; set; }
        public string BG { get; set; }
        public string LH { get; set; }
        public string BW { get; set; }
        public string BGR_UP { get; set; }
        public string BGR_DOWN { get; set; }
        public string FC { get; set; }
        public string DG_A { get; set; }
        public string DG_B { get; set; }
        public string SCH { get; set; }
        public string Z1 { get; set; }
        public string Z2 { get; set; }
        public string PA { get; set; }
        public string AH { get; set; }
        public string S1_level { get; set; }
        public string bar_code { get; set; }
        public string qr_code { get; set; }
        public string pro_code { get; set; }
        public string LG { get; set; }
        public string GS { get; set; }
        public string AS { get; set; }
        public string M { get; set; }
        public string N { get; set; }
        public string SHUL { get; set; }
        public string second_level { get; set; }
        public string pp_level { get; set; }
        public string A19 { get; set; }
        public string A01 { get; set; }

        public OME_QC GetOME_QC_data(DataTable dataTable)
        {
            Comm _comm = new Comm();
            OME_QC oME_QC =  new OME_QC()
            {
                TZ = _comm.Get_FloatFormat(_comm.Get_DataTableValue(dataTable, "TZ")),
                TW = _comm.Get_FloatFormat(_comm.Get_DataTableValue(dataTable, "TW")),
                DZ = _comm.Get_FloatFormat(_comm.Get_DataTableValue(dataTable, "DZ")),
                DW = _comm.Get_FloatFormat(_comm.Get_DataTableValue(dataTable, "DW")),
                W1 = _comm.Get_FloatFormat(_comm.Get_DataTableValue(dataTable, "W1")),
                W2 = _comm.Get_FloatFormat(_comm.Get_DataTableValue(dataTable, "W2")),
                BG = _comm.Get_FloatFormat(_comm.Get_DataTableValue(dataTable, "BG")),
                LH = _comm.Get_FloatFormat(_comm.Get_DataTableValue(dataTable, "LH")),
                BW = _comm.Get_FloatFormat(_comm.Get_DataTableValue(dataTable, "BW")),
                PA = _comm.Get_FloatFormat(_comm.Get_DataTableValue(dataTable, "PA")),
                AH = _comm.Get_FloatFormat(_comm.Get_DataTableValue(dataTable, "AH")),
                BGR_UP = _comm.Get_DataTableValue(dataTable, "BGR_UP"),
                BGR_DOWN = _comm.Get_DataTableValue(dataTable, "BGR_DOWN"),
                FC = _comm.Get_DataTableValue(dataTable, "FC"),
                DG_A = _comm.Get_DataTableValue(dataTable, "DG_A"),
                DG_B = _comm.Get_DataTableValue(dataTable, "DG_B"),
                SCH = _comm.Get_DataTableValue(dataTable, "SCH"),
                Z1 = _comm.Get_DataTableValue(dataTable, "Z1"),
                Z2 = _comm.Get_DataTableValue(dataTable, "Z2"),
                S1_level = _comm.Get_DataTableValue(dataTable, "S1_level"),
                bar_code = _comm.Get_DataTableValue(dataTable, "bar_code"),
                qr_code = _comm.Get_DataTableValue(dataTable, "qr_code"),
                pro_code = _comm.Get_DataTableValue(dataTable, "pro_code"),
                LG = _comm.Get_DataTableValue(dataTable, "LG"),
                GS = _comm.Get_DataTableValue(dataTable, "GS"),
                AS = _comm.Get_DataTableValue(dataTable, "AS"),
                M = _comm.Get_DataTableValue(dataTable, "M"),
                N = _comm.Get_DataTableValue(dataTable, "N"),
                SHUL = _comm.Get_DataTableValue(dataTable, "SHUL"),
                second_level = _comm.Get_DataTableValue(dataTable, "second_level"),
                pp_level = _comm.Get_DataTableValue(dataTable, "pp_level"),
                A19 = _comm.Get_DataTableValue(dataTable, "A19"),
                A01 = _comm.Get_DataTableValue(dataTable, "A01")
            };
            return oME_QC;
        }


    }

  
}
