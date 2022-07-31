using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace PRT_DTS.PrintData
{
    public class Print
    {
        public Print _print { get; set; }


        [DisplayName("印表機型號")]
        public string PrinerName { get; set; }
        [DisplayName("速度")]
        public decimal Speed { get; set; } = 3;
        [DisplayName("濃度")]
        public decimal DENSITY { get; set; } = 15;
        [DisplayName("大小寬度")]
        public decimal SizeX { get; set; } = 60;
        [DisplayName("大小高度")]
        public decimal SizeY { get; set; } = 75;
        [DisplayName("單位")]
        public string SizeUnit { get; set; } = "mm";
        [DisplayName("間距寬度")]
        public decimal GAPX { get; set; } = new Comm().String_ParseDecimal("1.5");
        [DisplayName("間距高度")]
        public decimal GAPY { get; set; } = new Comm().String_ParseDecimal("1.5");
        [DisplayName("間距單位")]
        public string GAPUnit { get; set; } = "mm";
        [DisplayName("相對起點")]
        public decimal DERECTION { get; set; } = 1;
        [DisplayName("偏移邊框寬度")]
        public decimal REFERENCEX { get; set; } = 0;
        [DisplayName("偏移邊框高度")]
        public decimal REFERENCEY { get; set; } = 0;
        [DisplayName("偏移邊框單位")]
        public string REFERENCEUnit { get; set; } = "mm";


        public class ProFormat
        {
            public int id { get; set; }
            public string no { get; set; }
            public string pro_code { get; set; }
            public string pro_name { get; set; }
            public string pro_qty { get; set; }
            public string lot_no { get; set; }
        }
    }
}
