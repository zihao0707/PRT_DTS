using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Net.NetworkInformation;
using System.Runtime.InteropServices;
using System.Text;

namespace PRT_DTS
{
    class Comm
    {
        public Comm()
        {
            IniFile ini = new IniFile("./setting.ini");
            //讀取 Ini 中 Num 的值 
            _DbConntion_ERP = ini.ReadIni("config", "MES");
            //_DbConntion_SCM = ini.ReadIni("config", "WMS");
            //_DbConntion_ERP = "Data Source=218.35.166.35;Initial Catalog=WMS_CloudPOC15_RFSOFT;Persist Security Info=True;User ID=sa;Password=1208jsh";
            //_DbConntion_SCM = "Data Source=218.35.166.35;Initial Catalog=SCM_CloudDEV;Persist Security Info=True;User ID=sa;Password=1208jsh";
        }
        public string _DbConntion_ERP { get; set; }
        public string _DbConntion_SCM { get; set; }

        /// <summary>
        /// 與資料庫做連線
        /// </summary>
        /// <returns>回傳一個SqlConnection的物件</returns>
        public SqlConnection Set_DBConnection(string DBname)
        {
            string connection = "";
            switch (DBname)
            {
                case "MES":
                    connection = Value_解密(_DbConntion_SCM);
                    break;
                case "SCM":
                    connection = Value_解密(_DbConntion_ERP);
                    break;
            }

            SqlConnection Connection_Db = new SqlConnection(connection);
            return Connection_Db;

        }

        /// <summary>
        /// 傳入一個SQL語法，回傳一個DataTable
        /// </summary>
        /// <param name="pSql">Select語法</param>
        /// <returns></returns>
        public DataTable Get_DataTable(string pSql, string sConnect)
        {
            DataTable datatable = new DataTable();
            try
            {
                if (pSql.Length > 0)
                {
                    using SqlConnection con_db = Set_DBConnection(sConnect);
                    con_db.Open();
                    SqlDataAdapter Adapter = new SqlDataAdapter(pSql, con_db);
                    Adapter.Fill(datatable);
                    con_db.Close();
                }
                return datatable;
            }
            catch (Exception) { throw; }
        }

        /// <summary>
        /// 執行sql 語法
        /// </summary>
        /// <param name="pSql"></param>
        /// <param name="connection"></param>
        public void SaveSQL(string pSql, string sConnect)
        {
            if (pSql != "")
            {
                try
                {
                    SqlConnection con_db = Set_DBConnection(sConnect);
                    con_db.Open();
                    SqlCommand cmd = new SqlCommand(pSql);
                    cmd.ExecuteNonQuery();
                    con_db.Close();
                }
                catch { }
            }
        }

        /// <summary>
        ///測試與資料庫做連線
        /// </summary>
        /// <returns>回傳一個SqlConnection的物件</returns>
        public bool Try_DBConnection(string sConnect)
        {
            try
            {
                using SqlConnection con_db = Set_DBConnection(sConnect);
                con_db.Open();
                con_db.Close();
                return true;
            }
            catch { return false; }
        }

        public List<string> GetMACAddress()
        {
            List<string> GetMACA = new List<string>();
            NetworkInterface[] nics = NetworkInterface.GetAllNetworkInterfaces();

            foreach (NetworkInterface adapter in nics)
            {
                IPInterfaceProperties properties = adapter.GetIPProperties();
                GetMACA.Add(adapter.GetPhysicalAddress().ToString());
            }
            return GetMACA;
        }


        private string Value_解密(string value)
        {
            string NewValue = "";

            for (int i = 1; i <= value.Length; i++)
            {
                if (i % 2 == 0)
                { NewValue += Value_解密處理區1(value.Substring(i - 1, 1)); }
                else
                { NewValue += Value_解密處理區2(value.Substring(i - 1, 1)); }

            }
            return NewValue;
        }

        private string Value_解密處理區1(string NewValue)
        {
            switch (NewValue)
            {
                case "a": NewValue = "!"; break;
                case "1": NewValue = "#"; break;
                case "m": NewValue = "$"; break;
                case "H": NewValue = "%"; break;
                case "8": NewValue = "&"; break;
                case "g": NewValue = "("; break;
                case "o": NewValue = ")"; break;
                case "0": NewValue = "*"; break;
                case "O": NewValue = "+"; break;
                case "k": NewValue = ","; break;
                case "K": NewValue = "-"; break;
                case "d": NewValue = "."; break;
                case "x": NewValue = "/"; break;
                case "R": NewValue = "0"; break;
                case "{": NewValue = "1"; break;
                case ",": NewValue = "2"; break;
                case "<": NewValue = "3"; break;
                case "Y": NewValue = "4"; break;
                case "q": NewValue = "5"; break;
                case "A": NewValue = "6"; break;
                case "(": NewValue = "7"; break;
                case "/": NewValue = "8"; break;
                case "&": NewValue = "9"; break;
                case "I": NewValue = ":"; break;
                case "J": NewValue = ";"; break;
                case "e": NewValue = "<"; break;
                case "u": NewValue = "="; break;
                case "6": NewValue = ">"; break;
                case "w": NewValue = "?"; break;
                case "D": NewValue = "@"; break;
                case "#": NewValue = "A"; break;
                case "c": NewValue = "B"; break;
                case "j": NewValue = "C"; break;
                case "!": NewValue = "D"; break;
                case ">": NewValue = "E"; break;
                case "y": NewValue = "F"; break;
                case "T": NewValue = "G"; break;
                case "~": NewValue = "H"; break;
                case "i": NewValue = "I"; break;
                case "F": NewValue = "J"; break;
                case ")": NewValue = "K"; break;
                case "b": NewValue = "L"; break;
                case "h": NewValue = "M"; break;
                case "%": NewValue = "N"; break;
                case "-": NewValue = "O"; break;
                case "t": NewValue = "P"; break;
                case "s": NewValue = "Q"; break;
                case "X": NewValue = "R"; break;
                case "f": NewValue = "S"; break;
                case "*": NewValue = "T"; break;
                case "l": NewValue = "U"; break;
                case "_": NewValue = "V"; break;
                case "^": NewValue = "W"; break;
                case "n": NewValue = "X"; break;
                case "+": NewValue = "Y"; break;
                case "P": NewValue = "Z"; break;
                case "U": NewValue = "["; break;
                case "W": NewValue = "\\"; break;
                case "r": NewValue = "]"; break;
                case "z": NewValue = "^"; break;
                case "N": NewValue = "_"; break;
                case "p": NewValue = "`"; break;
                case "`": NewValue = "a"; break;
                case "$": NewValue = "b"; break;
                case "M": NewValue = "c"; break;
                case "2": NewValue = "d"; break;
                case "S": NewValue = "e"; break;
                case "v": NewValue = "f"; break;
                case "4": NewValue = "g"; break;
                case ".": NewValue = "h"; break;
                case ";": NewValue = "i"; break;
                case " ": NewValue = "j"; break;
                case "9": NewValue = "k"; break;
                case "C": NewValue = "l"; break;
                case "B": NewValue = "m"; break;
                case "}": NewValue = "n"; break;
                case "\\": NewValue = "o"; break;
                case "G": NewValue = "p"; break;
                case "V": NewValue = "q"; break;
                case "E": NewValue = "r"; break;
                case "L": NewValue = "s"; break;
                case "3": NewValue = "t"; break;
                case "?": NewValue = "u"; break;
                case ":": NewValue = "v"; break;
                case "@": NewValue = "w"; break;
                case "5": NewValue = "x"; break;
                case "]": NewValue = "y"; break;
                case "7": NewValue = "z"; break;
                case "[": NewValue = "{"; break;
                case "=": NewValue = "}"; break;
                case "Q": NewValue = "~"; break;
                case "Z": NewValue = " "; break;
            }
            return NewValue;
        }

        private string Value_解密處理區2(string NewValue)
        {
            switch (NewValue)
            {
                case "n": NewValue = "!"; break;
                case "F": NewValue = "#"; break;
                case "p": NewValue = "$"; break;
                case "A": NewValue = "%"; break;
                case "w": NewValue = "&"; break;
                case "1": NewValue = "("; break;
                case "]": NewValue = ")"; break;
                case "f": NewValue = "*"; break;
                case "E": NewValue = "+"; break;
                case "5": NewValue = ","; break;
                case "L": NewValue = "-"; break;
                case "r": NewValue = "."; break;
                case "T": NewValue = "/"; break;
                case "i": NewValue = "0"; break;
                case "H": NewValue = "1"; break;
                case "<": NewValue = "2"; break;
                case "@": NewValue = "3"; break;
                case "{": NewValue = "4"; break;
                case "c": NewValue = "5"; break;
                case "B": NewValue = "6"; break;
                case "k": NewValue = "7"; break;
                case "^": NewValue = "8"; break;
                case "y": NewValue = "9"; break;
                case "W": NewValue = ":"; break;
                case "+": NewValue = ";"; break;
                case "6": NewValue = "<"; break;
                case "u": NewValue = "="; break;
                case "J": NewValue = ">"; break;
                case "C": NewValue = "?"; break;
                case "P": NewValue = "@"; break;
                case "t": NewValue = "A"; break;
                case "}": NewValue = "B"; break;
                case "[": NewValue = "C"; break;
                case "K": NewValue = "D"; break;
                case "?": NewValue = "E"; break;
                case "e": NewValue = "F"; break;
                case "j": NewValue = "G"; break;
                case "U": NewValue = "H"; break;
                case "2": NewValue = "I"; break;
                case ">": NewValue = "J"; break;
                case "a": NewValue = "K"; break;
                case "=": NewValue = "L"; break;
                case "o": NewValue = "M"; break;
                case "9": NewValue = "N"; break;
                case ".": NewValue = "O"; break;
                case "(": NewValue = "P"; break;
                case "_": NewValue = "Q"; break;
                case "v": NewValue = "R"; break;
                case "3": NewValue = "S"; break;
                case "!": NewValue = "T"; break;
                case "h": NewValue = "U"; break;
                case "~": NewValue = "V"; break;
                case "b": NewValue = "W"; break;
                case "`": NewValue = "X"; break;
                case "7": NewValue = "Y"; break;
                case "*": NewValue = "Z"; break;
                case "G": NewValue = "["; break;
                case "l": NewValue = "\\"; break;
                case "R": NewValue = "]"; break;
                case "m": NewValue = "^"; break;
                case "s": NewValue = "_"; break;
                case "0": NewValue = "`"; break;
                case "-": NewValue = "a"; break;
                case "\\": NewValue = "b"; break;
                case "Z": NewValue = "c"; break;
                case ":": NewValue = "d"; break;
                case "#": NewValue = "e"; break;
                case "I": NewValue = "f"; break;
                case "D": NewValue = "g"; break;
                case ",": NewValue = "h"; break;
                case "N": NewValue = "i"; break;
                case "O": NewValue = "j"; break;
                case ";": NewValue = "k"; break;
                case "%": NewValue = "l"; break;
                case "Q": NewValue = "m"; break;
                case "z": NewValue = "n"; break;
                case "&": NewValue = "o"; break;
                case "x": NewValue = "p"; break;
                case "S": NewValue = "q"; break;
                case ")": NewValue = "r"; break;
                case "Y": NewValue = "s"; break;
                case "4": NewValue = "t"; break;
                case "M": NewValue = "u"; break;
                case "/": NewValue = "v"; break;
                case "8": NewValue = "w"; break;
                case "g": NewValue = "x"; break;
                case "X": NewValue = "y"; break;
                case "V": NewValue = "z"; break;
                case "d": NewValue = "{"; break;
                case "q": NewValue = "}"; break;
                case " ": NewValue = "~"; break;
                case "$": NewValue = " "; break;
            }
            return NewValue;
        }

        /// <summary>
        /// 取得小數點位數
        /// </summary>
        /// <param name="num">數值</param>
        /// <param name="digits">小數點位數</param>
        /// <returns></returns>
        public string Get_FloatFormat(string num, int digits = 3)
        {
            decimal oDecimal;
            if (!num.Contains('.')) { num += ".0"; }
            if (!decimal.TryParse(num, out oDecimal)) { return num; }
            decimal before = decimal.Parse(num);

            return decimal.Round(before, digits).ToString();
        }
        public string ParseString(object pValue) => pValue.ToString() ?? "";
        public Int32 String_ParseInt32(string pValue) => Int32.TryParse(pValue, out Int32 result) ? result : 0;
        public Int64 String_ParseInt64(string pValue) => Int64.TryParse(pValue, out Int64 result) ? result : 0;
        public decimal String_ParseDecimal(string pValue) => decimal.TryParse(pValue, out decimal result) ? result : 0;
        public double String_ParseDouble(string pValue) => double.TryParse(pValue, out double result) ? result : 0;
        public float String_ParseFloat(string pValue) => float.TryParse(pValue, out float result) ? result : 0;
        public bool String_ParseBoolean(string pValue) => bool.TryParse(pValue, out bool result) ? result : false;
        public DateTime String_ParseDateTime(string pValue) => DateTime.TryParse(pValue, out DateTime result) ? result : DateTime.Now;
        public string Get_DateToString(string pValue) => String_ParseDateTime(pValue).ToString("yyyy-MM-dd");
        public string Get_TimeToString(string pValue) => String_ParseDateTime(pValue).ToString("HH:mm:ss");
        public string Get_DateTimeNow(string Format = "yyyy/MM/dd HH:mm:ss") => DateTime.Now.ToString(Format);
    }

    /// <summary>
    /// .ini檔處理
    /// </summary>
    /// <param name="section">標題</param>
    /// <param name="key">指定內容</param>
    /// <returns></returns>
    public class IniFile
    {

        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);

        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);

        private string filepath;
        public IniFile(string filepath)
        {
            this.filepath = filepath;
        }

        public string ReadIni(string section, string key)
        {
            StringBuilder temp = new StringBuilder(255);
            GetPrivateProfileString(section, key, "", temp, 255, filepath);
            return temp.ToString();
        }

    }
}
