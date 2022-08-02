using PRT_DTS.PrintData;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
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
            _DbConntion_MES = ini.ReadIni("config", "MES");
            //_DbConntion_SCM = ini.ReadIni("config", "WMS");
            //_DbConntion_ERP = "Data Source=218.35.166.35;Initial Catalog=WMS_CloudPOC15_RFSOFT;Persist Security Info=True;User ID=sa;Password=1208jsh";
            //_DbConntion_SCM = "Data Source=218.35.166.35;Initial Catalog=SCM_CloudDEV;Persist Security Info=True;User ID=sa;Password=1208jsh";
        }
        public string _DbConntion_MES { get; set; }
        public string _DbConntion_WMS { get; set; }

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
                    connection = Value_解密(_DbConntion_MES);
                    break;
                    //case "WMS":
                    //    connection = Value_解密(_DbConntion_WMS);
                    //    break;
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
            catch (Exception ex) { throw; }
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
                catch (Exception ex) { }
            }
        }

        public bool Insert_DB_Table(PRT02_0000 pData, string sTable, string sConnect)
        {

            if (pData.prt02_0000.Any())
            {

                using SqlConnection con_db = Set_DBConnection(sConnect);
                con_db.Open();
                SqlBulkCopy bulk = new SqlBulkCopy(con_db);
                bulk.DestinationTableName = sTable;
                try
                {
                    bulk.WriteToServer((IDataReader)pData);
                    con_db.Close();
                }
                catch (Exception ex)
                {
                    //錯誤處理 throw;
                    return false;
                }
            }
            return true;
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

        //取得配對碼
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

        //將DataTable轉換為string
        //searchName 要從DataTable取出的欄位
        public string Get_DataTableValue(DataTable data, string searchName)
        {
            List<string> colItem = data.Columns.Cast<DataColumn>().Select(x => x.ColumnName).ToList();
            string colName = searchName.Trim();
            return colItem.Where(x => x == colName).Any() ?
                data.AsEnumerable().Select(row => row[colName].ToString()).FirstOrDefault()
                : "";
        }

        //將DataTable轉換為list
        //searchName 要從DataTable取出的欄位
        public List<string> Get_DataTableValue_list(DataTable data, string searchName)
        {
            List<string> colItem = data.Columns.Cast<DataColumn>().Select(x => x.ColumnName).ToList();
            string colName = searchName.Trim();
            List<string> list = colItem.Where(x => x == colName).Any() ?
                data.AsEnumerable().Select(row => row[colName].ToString()).ToList()
                : new List<string>();
            return list;
        }

        /// <summary>
        /// 執行單一刪除語法
        /// </summary>
        /// <param name="pTableCode">執行刪除的目標Table</param>
        /// <param name="pKeyCode">鍵值欄位</param>
        /// <param name="pKeyValue">鍵值</param>
        public void Del_QueryData(string pTableCode, string pKeyCode, string pKeyValue, string sConnect)
        {
            string sSql = " DELETE " + pTableCode +
                          "  WHERE " + pKeyCode + "   = @" + pKeyCode;
            using (SqlConnection con_db = Set_DBConnection(sConnect))
            {
                con_db.Open();
                SqlCommand sqlCommand = new SqlCommand(sSql);
                sqlCommand.Connection = con_db;
                sqlCommand.Parameters.Add(new SqlParameter(pKeyCode, pKeyValue));
                sqlCommand.ExecuteNonQuery();
                con_db.Close();
            }
        }




        /// <summary>
        /// 取得新的GUID
        /// </summary>
        /// <param name="Type">
        /// D : 36 個字符 (同等 "")
        /// N : 32 個字符
        /// B : 38 個字符（大括號）
        /// P : 38 個字符（括號）
        /// X : 68 個字符（十六進制）
        /// </param>
        /// <returns></returns>
        public string Get_NewGUID(string Type = "N")
        {
            return Guid.NewGuid().ToString(Type);
        }

        //組合流水號 (範例格式'R' + yymmdd + 0000)
        //rules 流水號格式
        //index 從序號幾開始
        public string Get_RulesKey(string rules, int index)
        {
            List<string> Item = rules.Split('+').ToList();
            string result = "";
            foreach (string str in Item)
            {
                if (str.Contains("'"))
                {
                    result += str.Replace("'", "");
                    continue;
                }

                int outResult;
                if (int.TryParse(str, out outResult))
                {
                    string Key = Get_ScrNo(index, str);
                    result += Key;
                    continue;
                }

                string date = DateTime.Now.ToString(str);
                result += date;
            }
            return result;
        }
        //組合流水序號數字編號部分
        public string Get_ScrNo(int key, string str)
        {
            int range = str.Length - key.ToString().Length;
            char[] array = str.ToArray();
            int index = 0; string result = "";
            while (index < range)
            {
                result += array[index].ToString();
                index++;
            }
            result += key;
            return result;
        }

        public bool ins_PRT02_0000(PRT01_0000 prt, string sConnect)
        {

            PRT02_0000 prt02 = new PRT02_0000();
            prt02.prt02_0000 = new Comm().Get_NewGUID();
            prt02.prt_type = prt.PrtType;
            prt02.prt_kind = prt.PrtKind;
            prt02.print_name = prt.PrintName;
            prt02.print_data = prt.PrintData.Replace("'","''");
            prt02.result = "OK";
            prt02.usr_code = prt.UsrCode;
            prt02.ins_date = DateTime.Now.ToString("yyyy-MM-dd");
            prt02.ins_time = DateTime.Now.ToString("HH:mm:ss");
            prt02.prt_date = DateTime.Now.ToString("yyyy-MM-dd");
            prt02.prt_time = DateTime.Now.ToString("HH:mm:ss");

            string insSql = @$"INSERT INTO PRT02_0000 VALUES (
                            '{prt02.prt02_0000}','{ prt02.prt_type}','{prt02.prt_kind}',
                            '{prt02.print_name}','{prt02.print_data}',
                            '{prt02.result}','','{ prt02.ins_date}','{prt02.ins_time}',
                            '{prt02.usr_code}','{prt02.prt_date}','{prt02.prt_time}','','')";

            using (SqlConnection con_db = Set_DBConnection(sConnect))
            {
                con_db.Open();
                SqlCommand sqlCommand = new SqlCommand(insSql);
                sqlCommand.Connection = con_db;
                sqlCommand.ExecuteNonQuery();
                con_db.Close();
            }
            return true;
        }

        //取得該標籤歷史當天的紀錄
        //ins_date 列印日期
        //rules_code 流水號編碼原則
        //LabelCode 標籤種類
        //RunNumber 有幾筆資料，累加數量
        public int getPrt_cnt有編碼格式(string ins_date, string rules_code, string LabelCode, int RunNumber)
        {
            rules_code = rules_code.Replace("'", "");
            Comm comm = new Comm();
            string sSql = @$" DECLARE @Columns VARCHAR(MAX)
                                        DECLARE @JsonValue VARCHAR(MAX)

                                        SELECT @Columns = COALESCE(@Columns + ',' + 
				                                          REPLACE(REPLACE(print_data, '[' ,''), ']' ,''), 
				                                          REPLACE(REPLACE(print_data, '[' ,''), ']' ,'')) 
                                        from PRT02_0000 where ins_date = '{ins_date}' and (print_data like '%{LabelCode}%' 
                                                                and REPLACE(print_data, '''' ,'') like '%{rules_code}%') ;
                                        SET @JsonValue = '['+ @Columns+']'; -- 將Json 

                                        SELECT sum(prt_cnt) as prt_cnt FROM OPENJSON(@JsonValue)
                                        WITH(
	                                        prt_cnt int '$.prt_cnt'
                                        )";
            var data = comm.Get_DataTable(sSql, "MES");

            int Oldprt_cnt = comm.String_ParseInt32(Get_DataTableValue(data, "prt_cnt"));
            Oldprt_cnt = Oldprt_cnt + RunNumber;
            return Oldprt_cnt;
        }


        //取得該標籤歷史當天的紀錄(無特別輸入編碼)
        //ins_date 列印日期
        //rules_code 流水號編碼原則
        //LabelCode 標籤種類
        //RunNumber 有幾筆資料，累加數量
        public int getPrt_cnt無編碼格式(string ins_date, string LabelCode)
        {
            Comm comm = new Comm();
            string sSql = @$" DECLARE @Columns VARCHAR(MAX)
                                        DECLARE @JsonValue VARCHAR(MAX)

                                        SELECT @Columns = COALESCE(@Columns + ',' + 
				                                          REPLACE(REPLACE(print_data, '[' ,''), ']' ,''), 
				                                          REPLACE(REPLACE(print_data, '[' ,''), ']' ,'')) 
                                        from PRT02_0000 where ins_date = {ins_date} and print_data like '%{LabelCode}%' ;
                                        SET @JsonValue = '['+ @Columns+']'; -- 將Json 

                                        SELECT sum(prt_cnt) as prt_cnt FROM OPENJSON(@JsonValue)
                                        WITH(
	                                        prt_cnt int '$.prt_cnt'
                                        )";
            var data = comm.Get_DataTable(sSql, "MES");

            int Oldprt_cnt = comm.String_ParseInt32(Get_DataTableValue(data, "prt_cnt"));
            Oldprt_cnt = Oldprt_cnt + 1;
            return Oldprt_cnt;
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
