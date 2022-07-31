using System;
using System.Collections.Generic;
using System.Text;
using Newtonsoft.Json;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;
using System.IO;
using PRT_DTS.PrintData;

namespace PRT_DTS
{
    public class TSCLIB
    {
        #region TSCLIB_函數內容
        [DllImport("TSCLIB.dll", EntryPoint = "about")]
        public static extern int about();


        [DllImport("TSCLIB.dll", EntryPoint = "openport")]
        public static extern int openport(string printername);

        [DllImport("TSCLIB.dll", EntryPoint = "barcode")]
        public static extern int barcode(string x, string y, string type,
                    string height, string readable, string rotation,
                    string narrow, string wide, string code);

        [DllImport("TSCLIB.dll", EntryPoint = "clearbuffer")]
        public static extern int clearbuffer();

        [DllImport("TSCLIB.dll", EntryPoint = "closeport")]
        public static extern int closeport();

        [DllImport("TSCLIB.dll", EntryPoint = "downloadpcx")]
        public static extern int downloadpcx(string filename, string image_name);

        [DllImport("TSCLIB.dll", EntryPoint = "formfeed")]
        public static extern int formfeed();

        [DllImport("TSCLIB.dll", EntryPoint = "nobackfeed")]
        public static extern int nobackfeed();

        [DllImport("TSCLIB.dll", EntryPoint = "printerfont")]
        public static extern int printerfont(string x, string y, string fonttype,
                        string rotation, string xmul, string ymul,
                        string text);

        [DllImport("TSCLIB.dll", EntryPoint = "printlabel")]
        public static extern int printlabel(string set, string copy);

        [DllImport("TSCLIB.dll", EntryPoint = "sendcommand")]
        public static extern int sendcommand(string printercommand);

        [DllImport("TSCLIB.dll", EntryPoint = "setup")]
        public static extern int setup(string width, string height,
                  string speed, string density,
                  string sensor, string vertical,
                  string offset);

        [DllImport("TSCLIB.dll", EntryPoint = "windowsfont")]
        public static extern int windowsfont(int x, int y, int fontheight,
                        int rotation, int fontstyle, int fontunderline,
                        string szFaceName, string content);

        [DllImport("TSCLIB.dll", EntryPoint = "usbportqueryprinter")]
        public static extern int usbportqueryprinter();

        [DllImport("TSCLIB.dll", EntryPoint = "usbprintername")]
        public static extern int usbprintername();

        [DllImport("TSCLIB.dll", EntryPoint = "usbprinterserial")]
        public static extern int usbprinterserial();
        #endregion
        Comm _comm = new Comm();
        public void Command(string name, string txt)
        {
            openport(name);
            sendcommand(txt);
            closeport();
        }

        public void Print_LabelsData(PRT01_0000 pRT)
        {
            DataTable dataTable = JsonConvert.DeserializeObject<DataTable>(pRT.PrintData);
            openport(pRT.PrintName);//印表機名稱
            clearbuffer();

            switch (pRT.LabelCode)
            {
                case "FOMAT"://量治具標籤作業
                    List<string> QrCode = Get_DataTableValue_list(dataTable, "qrcode");
                    foreach (var value in QrCode)
                    {
                        FOMAT_Print(value);
                    }
                    break;

                case "IMG"://展會圖片

                    IMG_Print();
                    BARCODE(dataTable, pRT.LabelCode);
                    break;

                case "SET"://把號

                    OME_SET sET = new OME_SET(); ;
                    string insDate = Get_DataTableValue(dataTable, "ins_date");
                    int prt_cnt = new Comm().String_ParseInt32(Get_DataTableValue(dataTable, "prt_cnt"));
                    for (int i = 1; i <= prt_cnt; i++)
                    {
                        int Oldprt_cnt = getPrt_cnt有編碼格式(insDate, Get_DataTableValue(dataTable, "rules_code"), pRT.LabelCode, i);

                        string Key = Get_RulesKey(Get_DataTableValue(dataTable, "rules_code"), Oldprt_cnt);
                        Key = Key.Replace(" ", "");
                        //格式'R' + yymmdd + 0000
                        SET_Print(Key, pRT.PrintName);
                    }
                    break;

                case "QC"://檢驗標籤
                    QC_Print(pRT.PrintName, new OME_QC()
                    {
                        TZ = _comm.Get_FloatFormat(Get_DataTableValue(dataTable, "TZ")),
                        TW = _comm.Get_FloatFormat(Get_DataTableValue(dataTable, "TW")),
                        DZ = _comm.Get_FloatFormat(Get_DataTableValue(dataTable, "DZ")),
                        DW = _comm.Get_FloatFormat(Get_DataTableValue(dataTable, "DW")),
                        W1 = _comm.Get_FloatFormat(Get_DataTableValue(dataTable, "W1")),
                        W2 = _comm.Get_FloatFormat(Get_DataTableValue(dataTable, "W2")),
                        BG = _comm.Get_FloatFormat(Get_DataTableValue(dataTable, "BG")),
                        LH = _comm.Get_FloatFormat(Get_DataTableValue(dataTable, "LH")),
                        BW = _comm.Get_FloatFormat(Get_DataTableValue(dataTable, "BW")),
                        PA = _comm.Get_FloatFormat(Get_DataTableValue(dataTable, "PA")),
                        AH = _comm.Get_FloatFormat(Get_DataTableValue(dataTable, "AH")),
                        BGR_UP = Get_DataTableValue(dataTable, "BGR_UP"),
                        BGR_DOWN = Get_DataTableValue(dataTable, "BGR_DOWN"),
                        FC = Get_DataTableValue(dataTable, "FC"),
                        DG_A = Get_DataTableValue(dataTable, "DG_A"),
                        DG_B = Get_DataTableValue(dataTable, "DG_B"),
                        SCH = Get_DataTableValue(dataTable, "SCH"),
                        Z1 = Get_DataTableValue(dataTable, "Z1"),
                        Z2 = Get_DataTableValue(dataTable, "Z2"),
                        S1_level = Get_DataTableValue(dataTable, "S1_level"),
                        bar_code = Get_DataTableValue(dataTable, "bar_code"),
                        qr_code = Get_DataTableValue(dataTable, "qr_code"),
                        pro_code = Get_DataTableValue(dataTable, "pro_code"),
                        LG = Get_DataTableValue(dataTable, "LG"),
                        GS = Get_DataTableValue(dataTable, "GS"),
                        AS = Get_DataTableValue(dataTable, "AS"),
                        M = Get_DataTableValue(dataTable, "M"),
                        N = Get_DataTableValue(dataTable, "N"),
                        SHUL = Get_DataTableValue(dataTable, "SHUL"),
                        second_level = Get_DataTableValue(dataTable, "second_level"),
                        pp_level = Get_DataTableValue(dataTable, "pp_level"),
                        A19 = Get_DataTableValue(dataTable, "A19"),
                        A01 = Get_DataTableValue(dataTable, "A01"),
                    });
                    break;
            }

            /*--------執行向上-----------*/
        }
        /*----------------測試區向下-----------------*/

        public void QC_Print(string print_name, OME_QC qC) //下LINQ查詢語法抓prt01
        {
            string QrCode = qC.qr_code;
            string BARCODE = qC.bar_code;
            string FormatCommand = @$"
                SIZE 101 mm, 20mm
                GAP 0,0
                DIRECTION 0
                CLS

              DIAGONAL 400,  100, 400, 230, 3
              DIAGONAL 580,  100, 580, 230, 3
              DIAGONAL 760,  100, 760, 230, 3
              DIAGONAL 940,  100, 940, 230, 3


                BARCODE 50,170, ""128"",40,1,0,2,2, ""{BARCODE}""
                QRCODE 1000, 90, L, 8, A, 0, M2, X150, J1, ""{QrCode}""

                TEXT 40,   40, ""0"", 0, 12, 12, 0, ""{qC.pro_code}""
                TEXT 40,   100, ""5"", 0,   2,   1, 0, ""{qC.second_level}""

                TEXT 410, 120, ""ROMAN.TTF"", 0, 8, 8, ""TZ""
                TEXT 480, 120, ""ROMAN.TTF"", 0, 8, 8, ""{qC.TZ}""
                TEXT 410, 160, ""ROMAN.TTF"", 0, 8, 8, ""TW""
                TEXT 480, 160, ""ROMAN.TTF"", 0, 8, 8, ""{qC.TW}""
                TEXT 410, 200, ""ROMAN.TTF"", 0, 8, 8, ""PA""
                TEXT 480, 200, ""ROMAN.TTF"", 0, 8, 8, ""{qC.PA}""

                TEXT 590, 120, ""ROMAN.TTF"", 0, 8, 8, ""DZ""
                TEXT 590, 160, ""ROMAN.TTF"", 0, 8, 8, ""DW""
                TEXT 590, 200, ""ROMAN.TTF"", 0, 8, 8, ""AH""
                TEXT 660, 120, ""ROMAN.TTF"", 0, 8, 8, ""{qC.DZ}""
                TEXT 660, 160, ""ROMAN.TTF"", 0, 8, 8, ""{qC.DW}""
                TEXT 660, 200, ""ROMAN.TTF"", 0, 8, 8, ""{qC.AH}""

                TEXT 920, 40,  ""ROMAN.TTF"", 0, 8, 8, ""{qC.A01}""
                TEXT 770, 78,  ""ROMAN.TTF"", 0, 8, 8, ""QC PASS""
                TEXT 920, 78,  ""ROMAN.TTF"", 0, 8, 8, ""{qC.pp_level}""
                      
                TEXT 770, 120, ""ROMAN.TTF"", 0, 8, 8, ""LH""
                TEXT 770, 160, ""ROMAN.TTF"", 0, 8, 8, ""LG""
                TEXT 770, 200, ""ROMAN.TTF"", 0, 8, 8, ""BW""
                TEXT 840, 120, ""ROMAN.TTF"", 0, 8, 8, ""{qC.LH}""
                TEXT 840, 160, ""ROMAN.TTF"", 0, 8, 8, ""{qC.LG}""
                TEXT 840, 200, ""ROMAN.TTF"", 0, 8, 8, ""{qC.BW}""
            ";

            sendcommand(FormatCommand);//设置相对起点
            printlabel("1", "1");
            closeport();
        }

        public void SET_Print(string rules_code, string PrinerName)
        {
            string QrCode = rules_code;

            string FormatCommand = @$"
                SIZE 60 mm, 75 mm
                GAP 0,0
                DIRECTION 1
                CLS
                QRCODE 430, 600, L, 10, A, 0, M2, X250, J5, ""{QrCode}""
                TEXT 280,720, ""0"", 0, 12, 12, ""{QrCode}""
        ";
            sendcommand(FormatCommand);//设置相对起点
            printlabel("1", "1");
            closeport();



        }

        public void FOMAT_Print(string value) //量至距
        {

            string FormatCommand = @$" SIZE 101 mm, 50mm
                GAP 0,0
                SPEED 1
                DIRECTION 1
                CLS
              DIAGONAL 800,  20, 800, 400, 3
              DIAGONAL 1100,  20, 1100, 400, 3

              DIAGONAL 800,  20, 1100, 20, 3
              DIAGONAL 800,  400, 1100, 400, 3

             QRCODE 865, 90, L, 8, A, 0, M2, X150, J1, ""{value}""

                TEXT 850,280, ""0"", 0, 12, 12, ""{value}""";

            sendcommand(FormatCommand);//设置相对起点
            printlabel("1", "1");
            closeport();
        }

        public void IMG_Print() //下LINQ查詢語法抓prt01
        {

            downloadpcx(@"C:\Users\howard.chu\Desktop\img\TEST1.pcx", "BMP.PCX");
            string FormatCommand = @$"
                SIZE 101 mm, 101mm
                GAP 0,0
                SPEED 1
                DIRECTION 1
                CLS
                PUTPCX 68,20,""BMP.PCX""
               ";

            sendcommand(FormatCommand);//设置相对起点
            printlabel("1", "1");
            closeport();
        }

        public void BARCODE(DataTable dataTable, string LabelCode) //下LINQ查詢語法抓prt01
        {
            string insDate = Get_DataTableValue(dataTable, "ins_date");
            int Oldprt_cnt = getPrt_cnt無編碼格式(insDate, LabelCode);
            string ReelID = Get_RulesKey("'6971TW'+yyyyMMdd+000+'JS291843A'", Oldprt_cnt);
            ReelID = ReelID.Replace(" ", "");

            string DateCode = Get_RulesKey("yyyyMMdd+0000", Oldprt_cnt);
            DateCode = DateCode.Replace(" ", "");

            string FormatCommand = @$"
                SIZE 101 mm, 101mm
                GAP 0,0
                SPEED 1
                DIRECTION 1
                CLS
                DIAGONAL  50,  20, 1150,  20, 4
                DIAGONAL   50,  20, 50,  1100, 4
                DIAGONAL  50,  1100, 1150,  1100, 4 
                DIAGONAL   1150,  20, 1150,  1100, 4

                TEXT 100, 90, ""1"", 0, 3, 3, ""Reel ID : ""
                     BARCODE 400, 60, ""128"",100,2,0,2,2,""{ReelID}""
                TEXT 100, 290, ""1"", 0, 3, 3, ""P/N : ""
                        BARCODE 400, 260, ""128"",100,2,0,3,2,""8323081112014""
                TEXT 100, 450, ""1"", 0, 3, 3, ""Description : ""
                        TEXT 300, 530, ""1"", 0, 3, 3, ""MES/WMS/SCM/APS(TAIWAN)""
                TEXT 100, 630, ""1"", 0, 3, 3, ""Quantity : ""
                        BARCODE 450, 600, ""128"",100,2,0,4,2,""1""
                TEXT 100, 820, ""1"", 0, 3, 3, ""Date Code : ""
                        BARCODE 450, 780, ""128"",100,2,0,3,2,""{DateCode}""
                TEXT 100, 990, ""1"", 0, 3, 3, ""Vender Code : ""
                        BARCODE 510, 950, ""128"",100,2,0,3,2,""75778288""
                TEXT 800, 950, ""2"", 0, 3, 3, ""METAiM""
               ";

            sendcommand(FormatCommand);//设置相对起点
            printlabel("1", "1");
            closeport();


        }

        public void Test_Print()
        {

            string Key = Get_RulesKey("'PUME'+yymmdd+'-'+0000", 1);
            string FormatCommand = @$"
                    QRCODE 430, 280, L, 8, A, 0, M2, X250, J5, ""{Key}""
                    TEXT 250,380, ""0"", 0, 12, 12, ""{Key}""
                    PRINT 1,1
                    CLS
                ";
            sendcommand(FormatCommand);//设置相对起点
            printlabel("1", "1");
            closeport();
        }


        private string Get_DataTableValue(DataTable data, string searchName)
        {
            List<string> colItem = data.Columns.Cast<DataColumn>().Select(x => x.ColumnName).ToList();
            string colName = searchName.Trim();
            return colItem.Where(x => x == colName).Any() ?
                data.AsEnumerable().Select(row => row[colName].ToString()).FirstOrDefault()
                : "";
        }

        private List<string> Get_DataTableValue_list(DataTable data, string searchName)
        {
            List<string> colItem = data.Columns.Cast<DataColumn>().Select(x => x.ColumnName).ToList();
            string colName = searchName.Trim();
            List<string> list = colItem.Where(x => x == colName).Any() ?
                data.AsEnumerable().Select(row => row[colName].ToString()).ToList()
                : new List<string>();
            return list;
        }


        private string Get_RulesKey(string rules, int index)
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

        #region Private Function

        private string Get_ScrNo(int key, string str)
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


        /// <summary>
        /// 讀取印表機資訊
        /// </summary>
        /// <param name="info"></param>
        private void Read_PrintInfo(Print info)
        {
            //setup(
            //    info.SizeX.ToString(), info.SizeY.ToString(),
            //    info.Speed.ToString(), info.DENSITY.ToString(),
            //    "0", info.GAPX.ToString(), info.GAPY.ToString());
            sendcommand(string.Format("SIZE {0},{1}", info.SizeX + " " + info.SizeUnit, info.SizeY + " " + info.SizeUnit));//设置条码大小
            sendcommand(string.Format("GAP {0},{1}", info.GAPX + " " + info.GAPUnit, info.GAPY + " " + info.GAPUnit));//设置条码间隙
            sendcommand(string.Format("SPEED {0}", info.Speed));//设置打印速度
            sendcommand(string.Format("DENSITY {0}", info.DENSITY));//设置墨汁浓度
            sendcommand(string.Format("DERECTION {0}", info.DERECTION));//设置相对起点
            sendcommand(string.Format("REFERENCE {0},{1}", info.REFERENCEX + " " + info.REFERENCEUnit, info.REFERENCEY + " " + info.REFERENCEUnit));//设置偏移边框
            clearbuffer();
        }

        /// <summary>
        /// 標準化範本
        /// </summary>
        private void Print_Template_v1_0()
        {
            string commandString = @"
                DIAGONAL 20, 20, 850, 20, 3
                DIAGONAL 20, 20, 20, 680, 3
                DIAGONAL 20, 680, 850, 680, 3
                DIAGONAL 850, 20, 850, 680, 3
                DIAGONAL 20, 83, 850, 83, 3
                DIAGONAL 20, 146, 850, 146, 3
                DIAGONAL 20, 209, 850, 209, 3
                DIAGONAL 20, 310, 850, 310, 3
                DIAGONAL 20, 373, 645, 373, 3
                DIAGONAL 20, 436, 475, 436, 3
                DIAGONAL 20, 499, 850, 499, 3
                DIAGONAL 20, 680, 850, 680, 3
                DIAGONAL 120, 20, 120, 499,3
                DIAGONAL 80, 499, 80, 680, 3
                DIAGONAL 475, 209, 475, 310, 3
                DIAGONAL 475, 373, 475, 499, 3
                DIAGONAL 535, 373, 535, 499, 3
                DIAGONAL 645, 310, 645, 499, 3
                DIAGONAL 705, 310, 705, 499, 3
                DIAGONAL 645, 499, 645, 680, 3
                DIAGONAL 850, 310, 705, 499, 3
                QRCODE 675,515,M,4,A,0,M2,S7, ""testQrCode""";
            sendcommand(commandString);
            windowsfont(30, 30, 45, 0, 2, 0, "微軟正黑體", "產品");
            windowsfont(30, 93, 45, 0, 2, 0, "微軟正黑體", "產名");
            windowsfont(30, 156, 45, 0, 2, 0, "微軟正黑體", "規格");
            windowsfont(30, 219, 45, 0, 2, 0, "微軟正黑體", "批號");
            windowsfont(30, 264, 45, 0, 2, 0, "微軟正黑體", "效期");
            windowsfont(30, 320, 45, 0, 2, 0, "微軟正黑體", "工單");
            windowsfont(30, 380, 45, 0, 2, 0, "微軟正黑體", "調配");
            windowsfont(30, 446, 45, 0, 2, 0, "微軟正黑體", "重量");
            windowsfont(30, 530, 45, 0, 2, 0, "微軟正黑體", "註  ");
            windowsfont(30, 590, 45, 0, 2, 0, "微軟正黑體", "記  ");
            windowsfont(485, 383, 45, 0, 2, 0, "微軟正黑體", "投  ");
            windowsfont(485, 440, 45, 0, 2, 0, "微軟正黑體", "序  ");
            windowsfont(655, 326, 45, 0, 2, 0, "微軟正黑體", "工  ");
            windowsfont(655, 383, 45, 0, 2, 0, "微軟正黑體", "單  ");
            windowsfont(655, 440, 45, 0, 2, 0, "微軟正黑體", "序  ");
        }

        private void Print_Template_v2_0()
        {
            string commandString = @"
                DIAGONAL  45,  30, 805,  30, 3
                DIAGONAL  45,  30,  45, 580, 3
                DIAGONAL  45, 580, 805, 580, 3
                DIAGONAL  45,  95, 805,  95, 3
                DIAGONAL  45, 180, 805, 180, 3
                DIAGONAL  45, 250, 805, 250, 3
                DIAGONAL  45, 335, 805, 335, 3
                DIAGONAL  45, 400, 587, 400, 3
                DIAGONAL  45, 462, 587, 462, 3
                DIAGONAL  45, 525, 587, 525, 3
                DIAGONAL 270,  30, 270, 580, 3
                DIAGONAL 587, 335, 587, 580, 3
                DIAGONAL 805,  30, 805, 580, 3
                QRCODE 615,370,M,5,A,0,M2,S7, ""testQrCode""";
            sendcommand(commandString);
        }
        private void Print_Template_v2_1()
        {
            string commandString = @"
                DIAGONAL  60,   40, 795,   40, 4
                DIAGONAL  60,   40,  60, 1174, 4
                DIAGONAL  60, 1174, 795, 1174, 4
                DIAGONAL  60,  220, 795,  220, 4
                DIAGONAL  60,  420, 795,  420, 4
                DIAGONAL  60,  600, 795,  600, 4
                DIAGONAL  60,  780, 795,  780, 4
                DIAGONAL  60,  960, 795,  960, 4
                DIAGONAL 300,   40, 300,  960, 4
                DIAGONAL 650,  600, 650,  780, 4
                DIAGONAL 795,   40, 795, 1174, 4
                QRCODE 320,980,H,4,B,0,M2,S7, ""testQrCode""";
            sendcommand(commandString);
        }

        public string Get_RulesValue(string rules)
        {
            string a = "";
            return a;
        }

        public int getPrt_cnt有編碼格式(string ins_date, string rules_code, string LabelCode, int RunNumber)
        {
            rules_code = rules_code.Replace("'", "");
            Comm comm = new Comm();
            string sSql = @$" DECLARE @Columns VARCHAR(MAX)
                                        DECLARE @JsonValue VARCHAR(MAX)

                                        SELECT @Columns = COALESCE(@Columns + ',' + 
				                                          REPLACE(REPLACE(print_data, '[' ,''), ']' ,''), 
				                                          REPLACE(REPLACE(print_data, '[' ,''), ']' ,'')) 
                                        from PRT02_0000 where ins_date = {ins_date} and (print_data like '%{LabelCode}%' 
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


        #endregion
    }
}

