using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using WolfApprove.Model.CustomClass;
using System.Web.Script.Serialization;
using System.Data;
using WolfApprove.Model.Extension;
using System.Xml.Linq;
using System.Globalization;
using System.Reflection.Emit;
using System.Text.RegularExpressions;

namespace JobAutoTemplate
{
    internal class Program
    {
        private static string dbConnectionStringWolf
        {
            get
            {
                var dbConnectionString = System.Configuration.ConfigurationSettings.AppSettings["dbConnectionStringWolf"];
                if (!string.IsNullOrEmpty(dbConnectionString))
                {
                    return dbConnectionString;
                }
                return "Integrated Security=SSPI;Initial Catalog=WolfApproveCore.tym-ypmt;Data Source=DESKTOP-SHLJS2M\\SQLEXPRESS2019;User Id=sa;Password=pass@word1;";
            }
        }
        static void Main(string[] args)
        {
            Console.WriteLine("---------- Start JobAutoTemplate at :: " + DateTime.Now + " ----------");
            WriteLogFile("---------- Start JobAutoTemplate at :: " + DateTime.Now + " ----------");
            if (!string.IsNullOrEmpty(DataTime_Run))
            {
                DateTime datenow = DateTime.Now.Date;

                List<DateTime> lisDate = new List<DateTime>();
                var culture = System.Globalization.CultureInfo.CurrentCulture; // ตามเครื่อง server

                var sdate = DataTime_Run.Split('|');
                foreach (var itemdate in sdate)
                {
                    if (DateTime.TryParseExact($"{itemdate}/{datenow.Year}", "MM/dd/yyyy",
                        culture, System.Globalization.DateTimeStyles.None, out DateTime parsedDate))
                    {
                        lisDate.Add(parsedDate);
                    }
                }
                // เช็คว่าวันที่ปัจจุบันอยู่ใน List หรือไม่
                if (lisDate.Contains(datenow))
                {
                    WriteLogFile("It's time to run the program.");
                    GetStartForm();
                }
                else
                {
                    WriteLogFile("It's not time to run the program yet.");
                }
            }
            WriteLogFile("---------- Finish JobAutoTemplate at :: " + DateTime.Now + " ----------");
            Console.WriteLine("---------- Finish JobAutoTemplate at :: " + DateTime.Now + " ----------");
        }
        public static void GetStartForm()
        {
            TRNMemo getmemoid = new TRNMemo();
            try
            {
                DbWolfDataContext dbWolf = new DbWolfDataContext(dbConnectionStringWolf);
                var GetmemoCom = dbWolf.TRNMemos.Where(x => x.TemplateId == 99 && x.StatusName == "Completed" && x.CompletedDate >= DateTime.Now.AddMonths(iIntervalTime)).OrderBy(t => t.CompletedDate).ToList();
                var trnmemoform = dbWolf.TRNMemoForms.Where(z => z.TemplateId == 99 && z.obj_label == "Area").ToList();
                List<TRNMemo> joinedData = (from memo in GetmemoCom
                                            join form in trnmemoform
                                            on memo.MemoId equals form.MemoId
                                            group memo by form.obj_value into grouped
                                            select grouped.OrderByDescending(m => m.ModifiedDate).FirstOrDefault()).ToList();

                Console.WriteLine("DCC Area : " + joinedData.Count());
                WriteLogFile("DCC Area : " + joinedData.Count());
                var Gettemid = dbWolf.MSTTemplates.Where(x => x.TemplateId == 101).FirstOrDefault();
                if (joinedData.Count > 0)
                {

                    string getadvance = Gettemid.AdvanceForm;
                    var jsonObject = JObject.Parse(getadvance);

                    var nonTableData = new List<KeyValuePair<string, string>>();
                    var tableData = new List<List<KeyValuePair<string, string>>>();
                    var tableTypeAndLabels = new List<KeyValuePair<string, string>>(); // สำหรับเก็บ 'type' และ 'label' ของ tables
                                                                                       // Accessing the 'items' array
                    if (jsonObject["items"] is JArray itemsArray)
                    {
                        foreach (var item in itemsArray)
                        {
                            // Accessing the 'layout' array within each 'item'
                            if (item["layout"] is JArray layoutArray)
                            {
                                foreach (var layout in layoutArray)
                                {
                                    var template = layout["template"];
                                    var type = (string)template["type"];

                                    // Check if the type is 'tb' (table)
                                    if (type == "tb")
                                    {
                                        var tableLabel = (string)template["label"];
                                        if (!string.IsNullOrEmpty(tableLabel))
                                        {
                                            tableTypeAndLabels.Add(new KeyValuePair<string, string>(type, tableLabel)); // เพิ่ม 'type' และ 'label' ของ table
                                        }

                                        var tableRow = new List<KeyValuePair<string, string>>();
                                        // Extract labels from the columns of the table
                                        if (template["attribute"]["column"] is JArray columns)
                                        {
                                            foreach (var column in columns)
                                            {
                                                var columnLabel = (string)column["label"];
                                                if (!string.IsNullOrEmpty(columnLabel))
                                                {
                                                    tableRow.Add(new KeyValuePair<string, string>(columnLabel, string.Empty));
                                                }
                                            }
                                            tableData.Add(tableRow);
                                        }
                                    }
                                    else
                                    {
                                        // For non-table types, just add the label and value
                                        var label = (string)template["label"];
                                        var value = (string)layout["data"]?["value"] ?? string.Empty;
                                        if (!string.IsNullOrEmpty(label))
                                        {
                                            nonTableData.Add(new KeyValuePair<string, string>(label, value));
                                        }
                                    }
                                }
                            }
                        }
                    }





                    foreach (var getdata in joinedData)
                    {
                        var tableRows = new Dictionary<string, List<KeyValuePair<string, string>>>();
                        Console.WriteLine("Ref Memoid : " + getdata.MemoId);
                        WriteLogFile("Ref Memoid : " + getdata.MemoId);
                        getmemoid.MemoId = getdata.MemoId;
                        string madvance = getdata.MAdvancveForm;
                        var jsonObject2 = JObject.Parse(madvance);

                        foreach (var item in jsonObject2["items"])
                        {
                            foreach (var layout in item["layout"])
                            {
                                var template = layout["template"];
                                var type = (string)template["type"];

                                if (type == "tb")
                                {
                                    var tableName = (string)template["label"];
                                    var columns = template["attribute"]["column"];
                                    var rows = layout["data"]["row"];

                                    int rowIndex = 0;
                                    foreach (var row in rows)
                                    {
                                        var rowKey = tableName + " - Row " + rowIndex;
                                        tableRows[rowKey] = new List<KeyValuePair<string, string>>();

                                        int columnIndex = 0;
                                        foreach (var columnValue in row)
                                        {
                                            var columnLabel = (string)columns[columnIndex]["label"];
                                            var value = (string)columnValue["value"];
                                            tableRows[rowKey].Add(new KeyValuePair<string, string>(columnLabel, value));
                                            columnIndex++;
                                        }
                                        rowIndex++;
                                    }
                                }
                                else
                                {
                                    // For other types, just get the value
                                    var label = (string)template["label"];

                                    var value = (string)layout["data"]["value"];
                                    if (!string.IsNullOrEmpty(label) && !string.IsNullOrEmpty(value))
                                    {
                                        var rowKey = "SingleField - " + label;
                                        tableRows[rowKey] = new List<KeyValuePair<string, string>> { new KeyValuePair<string, string>(label, value) };
                                    }
                                }
                            }
                        }


                        string nottbv = string.Empty;
                        string nottb = string.Empty;
                        foreach (var nontabelv in tableRows)
                        {
                            if (nontabelv.Value[0].Key == "Document No.")
                            {
                                WriteLogFile("Start running");
                                TRNControlRunning running = InsertDocumentNo(dbWolf);
                                WriteLogFile("running : " + running.RunningNumber);
                                getadvance = PushValue(getadvance, running.RunningNumber, nontabelv.Value[0].Key);
                            }
                            if (nontabelv.Value[0].Key == "ปี")
                            {
                                getadvance = PushValue(getadvance, nontabelv.Value[0].Value, nontabelv.Value[0].Key);
                            }
                            if (nontabelv.Value[0].Key == "ครั้งที่")
                            {
                                getadvance = PushValue(getadvance, nontabelv.Value[0].Value, nontabelv.Value[0].Key);
                            }
                            if (nontabelv.Value[0].Value == "01 : Management System" || nontabelv.Value[0].Value.Contains("01 :"))
                            {
                                if (nonTableData[4].Key == "ISO Area Code : 01")
                                {
                                    foreach (var nontabelv2 in tableRows)
                                    {
                                        if (nontabelv2.Key == "SingleField - Area")
                                        {
                                            nottbv = nontabelv.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottbv, nonTableData[5].Key);
                                        }

                                        if (nontabelv2.Key == "SingleField - Scope")
                                        {
                                            nottb = nontabelv2.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottb, nonTableData[6].Key);
                                        }
                                    }
                                }
                            }
                            if (nontabelv.Value[0].Value == "02 : Assembly Group" || nontabelv.Value[0].Value.Contains("02 :"))
                            {
                                if (nonTableData[7].Key == "ISO Area Code : 02")
                                {

                                    foreach (var nontabelv2 in tableRows)
                                    {
                                        if (nontabelv2.Key == "SingleField - Area")
                                        {
                                            nottbv = nontabelv.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottbv, nonTableData[8].Key);
                                        }

                                        if (nontabelv2.Key == "SingleField - Scope")
                                        {
                                            nottb = nontabelv2.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottb, nonTableData[9].Key);
                                        }
                                    }
                                }
                            }
                            if (nontabelv.Value[0].Value == "03 : Spray Painting Group" || nontabelv.Value[0].Value.Contains("03 :"))
                            {
                                if (nonTableData[10].Key == "ISO Area Code : 03")
                                {

                                    foreach (var nontabelv2 in tableRows)
                                    {
                                        if (nontabelv2.Key == "SingleField - Area")
                                        {
                                            nottbv = nontabelv.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottbv, nonTableData[11].Key);
                                        }

                                        if (nontabelv2.Key == "SingleField - Scope")
                                        {
                                            nottb = nontabelv2.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottb, nonTableData[12].Key);
                                        }
                                    }
                                }
                            }
                            if (nontabelv.Value[0].Value == "04 : Export/Import Group" || nontabelv.Value[0].Value.Contains("04 :"))
                            {
                                if (nonTableData[13].Key == "ISO Area Code : 04")
                                {

                                    foreach (var nontabelv2 in tableRows)
                                    {
                                        if (nontabelv2.Key == "SingleField - Area")
                                        {
                                            nottbv = nontabelv.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottbv, nonTableData[14].Key);
                                        }

                                        if (nontabelv2.Key == "SingleField - Scope")
                                        {
                                            nottb = nontabelv2.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottb, nonTableData[15].Key);
                                        }
                                    }
                                }
                            }
                            if (nontabelv.Value[0].Value == "05 : Maintenance Assembly Group" || nontabelv.Value[0].Value.Contains("05 :"))
                            {
                                if (nonTableData[16].Key == "ISO Area Code : 05")
                                {

                                    foreach (var nontabelv2 in tableRows)
                                    {
                                        if (nontabelv2.Key == "SingleField - Area")
                                        {
                                            nottbv = nontabelv.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottbv, nonTableData[17].Key);
                                        }

                                        if (nontabelv2.Key == "SingleField - Scope")
                                        {
                                            nottb = nontabelv2.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottb, nonTableData[18].Key);
                                        }
                                    }

                                }
                            }
                            if (nontabelv.Value[0].Value == "06 : Machining Group" || nontabelv.Value[0].Value.Contains("06 :"))
                            {
                                if (nonTableData[19].Key == "ISO Area Code : 06")
                                {

                                    foreach (var nontabelv2 in tableRows)
                                    {
                                        if (nontabelv2.Key == "SingleField - Area")
                                        {
                                            nottbv = nontabelv.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottbv, nonTableData[20].Key);
                                        }

                                        if (nontabelv2.Key == "SingleField - Scope")
                                        {
                                            nottb = nontabelv2.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottb, nonTableData[21].Key);
                                        }
                                    }
                                }
                            }
                            if (nontabelv.Value[0].Value == "07 : Welding / EDP & Sub Frame Group" || nontabelv.Value[0].Value.Contains("07 :"))
                            {
                                if (nonTableData[22].Key == "ISO Area Code : 07")
                                {

                                    foreach (var nontabelv2 in tableRows)
                                    {
                                        if (nontabelv2.Key == "SingleField - Area")
                                        {
                                            nottbv = nontabelv.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottbv, nonTableData[23].Key);
                                        }

                                        if (nontabelv2.Key == "SingleField - Scope")
                                        {
                                            nottb = nontabelv2.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottb, nonTableData[24].Key);
                                        }
                                    }
                                }
                            }
                            if (nontabelv.Value[0].Value == "08 : Maintenance Parts Group" || nontabelv.Value[0].Value.Contains("08 :"))
                            {
                                if (nonTableData[25].Key == "ISO Area Code : 08")
                                {

                                    foreach (var nontabelv2 in tableRows)
                                    {
                                        if (nontabelv2.Key == "SingleField - Area")
                                        {
                                            nottbv = nontabelv.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottbv, nonTableData[26].Key);
                                        }

                                        if (nontabelv2.Key == "SingleField - Scope")
                                        {
                                            nottb = nontabelv2.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottb, nonTableData[27].Key);
                                        }
                                    }
                                }
                            }
                            if (nontabelv.Value[0].Value == "09 : In Plant Parts Quality Assurance Group" || nontabelv.Value[0].Value.Contains("09 :"))
                            {
                                if (nonTableData[28].Key == "ISO Area Code : 09")
                                {

                                    foreach (var nontabelv2 in tableRows)
                                    {
                                        if (nontabelv2.Key == "SingleField - Area")
                                        {
                                            nottbv = nontabelv.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottbv, nonTableData[29].Key);
                                        }

                                        if (nontabelv2.Key == "SingleField - Scope")
                                        {
                                            nottb = nontabelv2.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottb, nonTableData[30].Key);
                                        }
                                    }
                                }
                            }
                            if (nontabelv.Value[0].Value == "10 : In Plant Parts Quality Control Group" || nontabelv.Value[0].Value.Contains("10 :"))
                            {
                                if (nonTableData[31].Key == "ISO Area Code : 10")
                                {

                                    foreach (var nontabelv2 in tableRows)
                                    {
                                        if (nontabelv2.Key == "SingleField - Area")
                                        {
                                            nottbv = nontabelv.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottbv, nonTableData[32].Key);
                                        }

                                        if (nontabelv2.Key == "SingleField - Scope")
                                        {
                                            nottb = nontabelv2.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottb, nonTableData[33].Key);
                                        }
                                    }
                                }
                            }
                            if (nontabelv.Value[0].Value == "11 : Production Support Center Group" || nontabelv.Value[0].Value.Contains("11 :"))
                            {
                                if (nonTableData[34].Key == "ISO Area Code : 11")
                                {

                                    foreach (var nontabelv2 in tableRows)
                                    {
                                        if (nontabelv2.Key == "SingleField - Area")
                                        {
                                            nottbv = nontabelv.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottbv, nonTableData[35].Key);
                                        }

                                        if (nontabelv2.Key == "SingleField - Scope")
                                        {
                                            nottb = nontabelv2.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottb, nonTableData[36].Key);
                                        }
                                    }
                                }
                            }
                            if (nontabelv.Value[0].Value == "12 : Production Development Group" || nontabelv.Value[0].Value.Contains("12 :"))
                            {
                                if (nonTableData[37].Key == "ISO Area Code : 12")
                                {

                                    foreach (var nontabelv2 in tableRows)
                                    {
                                        if (nontabelv2.Key == "SingleField - Area")
                                        {
                                            nottbv = nontabelv.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottbv, nonTableData[38].Key);
                                        }

                                        if (nontabelv2.Key == "SingleField - Scope")
                                        {
                                            nottb = nontabelv2.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottb, nonTableData[39].Key);
                                        }
                                    }
                                }
                            }
                            if (nontabelv.Value[0].Value == "13 : New Model Quality Assurance Group & Market Information Control Group & Quality Promotion & Planning Group & Technical Regulations Control Group" || nontabelv.Value[0].Value.Contains("13 :"))
                            {
                                if (nonTableData[40].Key == "ISO Area Code : 13")
                                {

                                    foreach (var nontabelv2 in tableRows)
                                    {
                                        if (nontabelv2.Key == "SingleField - Area")
                                        {
                                            nottbv = nontabelv.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottbv, nonTableData[41].Key);
                                        }

                                        if (nontabelv2.Key == "SingleField - Scope")
                                        {
                                            nottb = nontabelv2.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottb, nonTableData[42].Key);
                                        }
                                    }
                                }
                            }
                            if (nontabelv.Value[0].Value == "14 : Product Quality Improvement Group" || nontabelv.Value[0].Value.Contains("14 :"))
                            {
                                if (nonTableData[43].Key == "ISO Area Code : 14")
                                {

                                    foreach (var nontabelv2 in tableRows)
                                    {
                                        if (nontabelv2.Key == "SingleField - Area")
                                        {
                                            nottbv = nontabelv.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottbv, nonTableData[44].Key);
                                        }

                                        if (nontabelv2.Key == "SingleField - Scope")
                                        {
                                            nottb = nontabelv2.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottb, nonTableData[45].Key);
                                        }
                                    }
                                }
                            }
                            if (nontabelv.Value[0].Value == "15 : Production Planning & Control Division" || nontabelv.Value[0].Value.Contains("15 :"))
                            {
                                if (nonTableData[46].Key == "ISO Area Code : 15")
                                {

                                    foreach (var nontabelv2 in tableRows)
                                    {
                                        if (nontabelv2.Key == "SingleField - Area")
                                        {
                                            nottbv = nontabelv.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottbv, nonTableData[47].Key);
                                        }

                                        if (nontabelv2.Key == "SingleField - Scope")
                                        {
                                            nottb = nontabelv2.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottb, nonTableData[48].Key);
                                        }
                                    }
                                }
                            }
                            if (nontabelv.Value[0].Value == "16 : Procurement Division" || nontabelv.Value[0].Value.Contains("16 :"))
                            {
                                if (nonTableData[49].Key == "ISO Area Code : 16")
                                {

                                    foreach (var nontabelv2 in tableRows)
                                    {
                                        if (nontabelv2.Key == "SingleField - Area")
                                        {
                                            nottbv = nontabelv.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottbv, nonTableData[50].Key);
                                        }

                                        if (nontabelv2.Key == "SingleField - Scope")
                                        {
                                            nottb = nontabelv2.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottb, nonTableData[51].Key);
                                        }
                                    }
                                }
                            }
                            if (nontabelv.Value[0].Value == "17 : Assembly Engineering Division" || nontabelv.Value[0].Value.Contains("17 :"))
                            {
                                if (nonTableData[52].Key == "ISO Area Code : 17")
                                {

                                    foreach (var nontabelv2 in tableRows)
                                    {
                                        if (nontabelv2.Key == "SingleField - Area")
                                        {
                                            nottbv = nontabelv.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottbv, nonTableData[53].Key);
                                        }

                                        if (nontabelv2.Key == "SingleField - Scope")
                                        {
                                            nottb = nontabelv2.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottb, nonTableData[54].Key);
                                        }
                                    }
                                }
                            }
                            if (nontabelv.Value[0].Value == "18 : Parts Engineering Group" || nontabelv.Value[0].Value.Contains("18 :"))
                            {
                                if (nonTableData[55].Key == "ISO Area Code : 18")
                                {

                                    foreach (var nontabelv2 in tableRows)
                                    {
                                        if (nontabelv2.Key == "SingleField - Area")
                                        {
                                            nottbv = nontabelv.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottbv, nonTableData[56].Key);
                                        }

                                        if (nontabelv2.Key == "SingleField - Scope")
                                        {
                                            nottb = nontabelv2.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottb, nonTableData[57].Key);
                                        }
                                    }
                                }
                            }
                            if (nontabelv.Value[0].Value == "19 : General Affairs Division & Corporate Legal Group" || nontabelv.Value[0].Value.Contains("19 :"))
                            {
                                if (nonTableData[58].Key == "ISO Area Code : 19")
                                {

                                    foreach (var nontabelv2 in tableRows)
                                    {
                                        if (nontabelv2.Key == "SingleField - Area")
                                        {
                                            nottbv = nontabelv.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottbv, nonTableData[59].Key);
                                        }

                                        if (nontabelv2.Key == "SingleField - Scope")
                                        {
                                            nottb = nontabelv2.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottb, nonTableData[60].Key);
                                        }
                                    }
                                }
                            }
                            if (nontabelv.Value[0].Value == "20 : Human Resources Division" || nontabelv.Value[0].Value.Contains("20 :"))
                            {
                                if (nonTableData[61].Key == "ISO Area Code : 20")
                                {

                                    foreach (var nontabelv2 in tableRows)
                                    {
                                        if (nontabelv2.Key == "SingleField - Area")
                                        {
                                            nottbv = nontabelv.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottbv, nonTableData[62].Key);
                                        }

                                        if (nontabelv2.Key == "SingleField - Scope")
                                        {
                                            nottb = nontabelv2.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottb, nonTableData[63].Key);
                                        }
                                    }
                                }
                            }
                            if (nontabelv.Value[0].Value == "21 : Information Technology Division & Corporate Planning Division" || nontabelv.Value[0].Value.Contains("21 :") || nontabelv.Value[0].Value.Contains("21:"))
                            {
                                if (nonTableData[64].Key == "ISO Area Code : 21")
                                {

                                    foreach (var nontabelv2 in tableRows)
                                    {
                                        if (nontabelv2.Key == "SingleField - Area")
                                        {
                                            nottbv = nontabelv.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottbv, nonTableData[65].Key);
                                        }

                                        if (nontabelv2.Key == "SingleField - Scope")
                                        {
                                            nottb = nontabelv2.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottb, nonTableData[66].Key);
                                        }
                                    }
                                }
                            }
                            if (nontabelv.Value[0].Value == "22 : Finance & Credit Management Division &IA Group" || nontabelv.Value[0].Value.Contains("22 :"))
                            {
                                if (nonTableData[67].Key == "ISO Area Code : 22")
                                {

                                    foreach (var nontabelv2 in tableRows)
                                    {
                                        if (nontabelv2.Key == "SingleField - Area")
                                        {
                                            nottbv = nontabelv.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottbv, nonTableData[68].Key);
                                        }

                                        if (nontabelv2.Key == "SingleField - Scope")
                                        {
                                            nottb = nontabelv2.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottb, nonTableData[69].Key);
                                        }
                                    }
                                }
                            }
                            if (nontabelv.Value[0].Value == "23 : Accounting & Financial Planning Division" || nontabelv.Value[0].Value.Contains("23 :"))
                            {
                                if (nonTableData[70].Key == "ISO Area Code : 23")
                                {

                                    foreach (var nontabelv2 in tableRows)
                                    {
                                        if (nontabelv2.Key == "SingleField - Area")
                                        {
                                            nottbv = nontabelv.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottbv, nonTableData[71].Key);
                                        }

                                        if (nontabelv2.Key == "SingleField - Scope")
                                        {
                                            nottb = nontabelv2.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottb, nonTableData[72].Key);
                                        }
                                    }
                                }

                            }
                            if (nontabelv.Value[0].Value == "24 : Sales Administration Group & SCM Planning Group & Fleet & Finance Division" || nontabelv.Value[0].Value.Contains("24 :") || nontabelv.Value[0].Value.Contains("24:"))
                            {
                                if (nonTableData[73].Key == "ISO Area Code : 24")
                                {

                                    foreach (var nontabelv2 in tableRows)
                                    {
                                        if (nontabelv2.Key == "SingleField - Area")
                                        {
                                            nottbv = nontabelv.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottbv, nonTableData[74].Key);
                                        }

                                        if (nontabelv2.Key == "SingleField - Scope")
                                        {
                                            nottb = nontabelv2.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottb, nonTableData[75].Key);
                                        }
                                    }
                                }
                            }
                            if (nontabelv.Value[0].Value == "25 : Merchandise Operation & Engineering Group & Big Bike Business Division & Motor Sport Planning & activities Group & Racing Group & Customer Satisfaction Group" || nontabelv.Value[0].Value.Contains("25 :"))
                            {
                                if (nonTableData[76].Key == "ISO Area Code : 25")
                                {

                                    foreach (var nontabelv2 in tableRows)
                                    {
                                        if (nontabelv2.Key == "SingleField - Area")
                                        {
                                            nottbv = nontabelv.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottbv, nonTableData[77].Key);
                                        }

                                        if (nontabelv2.Key == "SingleField - Scope")
                                        {
                                            nottb = nontabelv2.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottb, nonTableData[78].Key);
                                        }
                                    }
                                }
                            }
                            if (nontabelv.Value[0].Value == "26 : Service Division" || nontabelv.Value[0].Value.Contains("26 :"))
                            {
                                if (nonTableData[79].Key == "ISO Area Code : 26")
                                {

                                    foreach (var nontabelv2 in tableRows)
                                    {
                                        if (nontabelv2.Key == "SingleField - Area")
                                        {
                                            nottbv = nontabelv.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottbv, nonTableData[80].Key);
                                        }

                                        if (nontabelv2.Key == "SingleField - Scope")
                                        {
                                            nottb = nontabelv2.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottb, nonTableData[81].Key);
                                        }
                                    }
                                }
                            }
                            if (nontabelv.Value[0].Value == "27 : Spare Parts Division" || nontabelv.Value[0].Value.Contains("27 :"))
                            {
                                if (nonTableData[82].Key == "ISO Area Code : 27")
                                {

                                    foreach (var nontabelv2 in tableRows)
                                    {
                                        if (nontabelv2.Key == "SingleField - Area")
                                        {
                                            nottbv = nontabelv.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottbv, nonTableData[83].Key);
                                        }

                                        if (nontabelv2.Key == "SingleField - Scope")
                                        {
                                            nottb = nontabelv2.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottb, nonTableData[84].Key);
                                        }
                                    }
                                }
                            }
                            if (nontabelv.Value[0].Value == "28 : Marketing Planning & Product Planning Division & Marine & Golf Car Group" || nontabelv.Value[0].Value.Contains("28 :"))
                            {
                                if (nonTableData[85].Key == "ISO Area Code : 28")
                                {

                                    foreach (var nontabelv2 in tableRows)
                                    {
                                        if (nontabelv2.Key == "SingleField - Area")
                                        {
                                            nottbv = nontabelv.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottbv, nonTableData[86].Key);
                                        }

                                        if (nontabelv2.Key == "SingleField - Scope")
                                        {
                                            nottb = nontabelv2.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottb, nonTableData[87].Key);
                                        }
                                    }
                                }
                            }
                            if (nontabelv.Value[0].Value == "29 : Marketing-AT& Marketing support Division & Marketing Sports & Moped Division" || nontabelv.Value[0].Value.Contains("29 :"))
                            {
                                if (nonTableData[88].Key == "ISO Area Code : 29")
                                {

                                    foreach (var nontabelv2 in tableRows)
                                    {
                                        if (nontabelv2.Key == "SingleField - Area")
                                        {
                                            nottbv = nontabelv.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottbv, nonTableData[89].Key);
                                        }

                                        if (nontabelv2.Key == "SingleField - Scope")
                                        {
                                            nottb = nontabelv2.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottb, nonTableData[90].Key);
                                        }
                                    }
                                }
                            }
                            if (nontabelv.Value[0].Value == "30 : Motor Sport Division" || nontabelv.Value[0].Value.Contains("30 :"))
                            {
                                if (nonTableData[91].Key == "ISO Area Code : 30")
                                {

                                    foreach (var nontabelv2 in tableRows)
                                    {
                                        if (nontabelv2.Key == "SingleField - Area")
                                        {
                                            nottbv = nontabelv.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottbv, nonTableData[92].Key);
                                        }

                                        if (nontabelv2.Key == "SingleField - Scope")
                                        {
                                            nottb = nontabelv2.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottb, nonTableData[93].Key);
                                        }
                                    }
                                }
                            }
                            if (nontabelv.Value[0].Value == "31 : Low Cost Automation & AI Group" || nontabelv.Value[0].Value.Contains("31 :"))
                            {
                                if (nonTableData[94].Key == "ISO Area Code : 31")
                                {

                                    foreach (var nontabelv2 in tableRows)
                                    {
                                        if (nontabelv2.Key == "SingleField - Area")
                                        {
                                            nottbv = nontabelv.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottbv, nonTableData[95].Key);
                                        }

                                        if (nontabelv2.Key == "SingleField - Scope")
                                        {
                                            nottb = nontabelv2.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottb, nonTableData[96].Key);
                                        }
                                    }
                                }
                            }
                            if (nontabelv.Value[0].Value == "32 : Logistics Innovation Division" || nontabelv.Value[0].Value.Contains("32 :"))
                            {
                                if (nonTableData[97].Key == "ISO Area Code : 32")
                                {

                                    foreach (var nontabelv2 in tableRows)
                                    {
                                        if (nontabelv2.Key == "SingleField - Area")
                                        {
                                            nottbv = nontabelv.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottbv, nonTableData[98].Key);
                                        }

                                        if (nontabelv2.Key == "SingleField - Scope")
                                        {
                                            nottb = nontabelv2.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottb, nonTableData[99].Key);
                                        }
                                    }
                                }
                            }
                            if (nontabelv.Value[0].Value == "33 : YRA Division" || nontabelv.Value[0].Value.Contains("33 :"))
                            {
                                if (nonTableData[100].Key == "ISO Area Code : 33")
                                {

                                    foreach (var nontabelv2 in tableRows)
                                    {
                                        if (nontabelv2.Key == "SingleField - Area")
                                        {
                                            nottbv = nontabelv.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottbv, nonTableData[101].Key);
                                        }

                                        if (nontabelv2.Key == "SingleField - Scope")
                                        {
                                            nottb = nontabelv2.Value[0].Value;
                                            getadvance = PushValue(getadvance, nottb, nonTableData[102].Key);
                                        }
                                    }
                                }
                            }
                        }


                        if (tableRows.Count > 0)
                        {
                            foreach (var tableRow in tableRows)
                            {
                                string rowKey = tableRow.Key;
                                List<KeyValuePair<string, string>> columns = tableRow.Value;
                                List<string> listvalue = new List<string>();
                                string keylable = string.Empty;
                                // Now iterate over the columns
                                foreach (var column in columns)
                                {
                                    if (column.Value == "01 : Management System" || column.Value.Contains("01 :"))
                                    {
                                        string valuelabeltb = tableTypeAndLabels[0].Value;
                                        foreach (var tablvalue in tableRows)
                                        {
                                            string rowKey2 = tablvalue.Key;
                                            List<KeyValuePair<string, string>> columnss = tablvalue.Value;
                                            foreach (var tablkey in columnss)
                                            {
                                                if (tablkey.Key == "Head working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "DCC Area")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "Working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "ID")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Name")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Position")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Division /Group / Section")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);
                                                                keylable = labelk.Key;

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }


                                            }
                                        }
                                        getadvance = ReplaceDataProcessBC(getadvance, listvalue, keylable, valuelabeltb);
                                        break;
                                    }
                                    if (column.Value == "02 : Assembly Group" || column.Value.Contains("02 :"))
                                    {
                                        string valuelabeltb = tableTypeAndLabels[1].Value;
                                        foreach (var tablvalue in tableRows)
                                        {
                                            string rowKey2 = tablvalue.Key;
                                            List<KeyValuePair<string, string>> columnss = tablvalue.Value;
                                            foreach (var tablkey in columnss)
                                            {
                                                if (tablkey.Key == "Head working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "DCC Area")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "Working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "ID")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Name")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Position")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Division /Group / Section")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);
                                                                keylable = labelk.Key;

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }


                                            }
                                        }
                                        getadvance = ReplaceDataProcessBC(getadvance, listvalue, keylable, valuelabeltb);
                                        break;
                                    }
                                    if (column.Value == "03 : Spray Painting Group" || column.Value.Contains("03 :"))
                                    {
                                        string valuelabeltb = tableTypeAndLabels[2].Value;
                                        foreach (var tablvalue in tableRows)
                                        {
                                            string rowKey2 = tablvalue.Key;
                                            List<KeyValuePair<string, string>> columnss = tablvalue.Value;
                                            foreach (var tablkey in columnss)
                                            {
                                                if (tablkey.Key == "Head working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "DCC Area")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "Working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "ID")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Name")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Position")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Division /Group / Section")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);
                                                                keylable = labelk.Key;

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }


                                            }
                                        }
                                        getadvance = ReplaceDataProcessBC(getadvance, listvalue, keylable, valuelabeltb);
                                        break;
                                    }
                                    if (column.Value == "04 : Export/Import Group" || column.Value.Contains("04 :"))
                                    {
                                        string valuelabeltb = tableTypeAndLabels[3].Value;
                                        foreach (var tablvalue in tableRows)
                                        {
                                            string rowKey2 = tablvalue.Key;
                                            List<KeyValuePair<string, string>> columnss = tablvalue.Value;
                                            foreach (var tablkey in columnss)
                                            {
                                                if (tablkey.Key == "Head working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "DCC Area")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "Working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "ID")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Name")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Position")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Division /Group / Section")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);
                                                                keylable = labelk.Key;

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }


                                            }
                                        }
                                        getadvance = ReplaceDataProcessBC(getadvance, listvalue, keylable, valuelabeltb);
                                        break;
                                    }
                                    if (column.Value == "05 : Maintenance Assembly Group" || column.Value.Contains("05 :"))
                                    {
                                        string valuelabeltb = tableTypeAndLabels[4].Value;
                                        foreach (var tablvalue in tableRows)
                                        {
                                            string rowKey2 = tablvalue.Key;
                                            List<KeyValuePair<string, string>> columnss = tablvalue.Value;
                                            foreach (var tablkey in columnss)
                                            {
                                                if (tablkey.Key == "Head working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "DCC Area")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "Working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "ID")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Name")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Position")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Division /Group / Section")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);
                                                                keylable = labelk.Key;

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }


                                            }
                                        }
                                        getadvance = ReplaceDataProcessBC(getadvance, listvalue, keylable, valuelabeltb);
                                        break;
                                    }
                                    if (column.Value == "06 : Machining Group" || column.Value.Contains("06 :"))
                                    {
                                        string valuelabeltb = tableTypeAndLabels[5].Value;
                                        foreach (var tablvalue in tableRows)
                                        {
                                            string rowKey2 = tablvalue.Key;
                                            List<KeyValuePair<string, string>> columnss = tablvalue.Value;
                                            foreach (var tablkey in columnss)
                                            {
                                                if (tablkey.Key == "Head working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "DCC Area")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "Working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "ID")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Name")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Position")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Division /Group / Section")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);
                                                                keylable = labelk.Key;

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }


                                            }
                                        }
                                        getadvance = ReplaceDataProcessBC(getadvance, listvalue, keylable, valuelabeltb);
                                        break;
                                    }
                                    if (column.Value == "07 : Welding / EDP & Sub Frame Group" || column.Value.Contains("07 :"))
                                    {
                                        string valuelabeltb = tableTypeAndLabels[6].Value;
                                        foreach (var tablvalue in tableRows)
                                        {
                                            string rowKey2 = tablvalue.Key;
                                            List<KeyValuePair<string, string>> columnss = tablvalue.Value;
                                            foreach (var tablkey in columnss)
                                            {
                                                if (tablkey.Key == "Head working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "DCC Area")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "Working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "ID")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Name")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Position")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Division /Group / Section")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);
                                                                keylable = labelk.Key;

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }


                                            }
                                        }
                                        getadvance = ReplaceDataProcessBC(getadvance, listvalue, keylable, valuelabeltb);
                                        break;
                                    }
                                    if (column.Value == "08 : Maintenance Parts Group" || column.Value.Contains("08 :"))
                                    {
                                        string valuelabeltb = tableTypeAndLabels[7].Value;
                                        foreach (var tablvalue in tableRows)
                                        {
                                            string rowKey2 = tablvalue.Key;
                                            List<KeyValuePair<string, string>> columnss = tablvalue.Value;
                                            foreach (var tablkey in columnss)
                                            {
                                                if (tablkey.Key == "Head working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "DCC Area")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "Working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "ID")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Name")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Position")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Division /Group / Section")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);
                                                                keylable = labelk.Key;

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }


                                            }
                                        }
                                        getadvance = ReplaceDataProcessBC(getadvance, listvalue, keylable, valuelabeltb);
                                        break;
                                    }
                                    if (column.Value == "09 : In Plant Parts Quality Assurance Group" || column.Value.Contains("09 :"))
                                    {
                                        string valuelabeltb = tableTypeAndLabels[8].Value;
                                        foreach (var tablvalue in tableRows)
                                        {
                                            string rowKey2 = tablvalue.Key;
                                            List<KeyValuePair<string, string>> columnss = tablvalue.Value;
                                            foreach (var tablkey in columnss)
                                            {
                                                if (tablkey.Key == "Head working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "DCC Area")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "Working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "ID")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Name")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Position")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Division /Group / Section")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);
                                                                keylable = labelk.Key;

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }


                                            }
                                        }
                                        getadvance = ReplaceDataProcessBC(getadvance, listvalue, keylable, valuelabeltb);
                                        break;
                                    }
                                    if (column.Value == "10 : In Plant Parts Quality Control Group" || column.Value.Contains("10 :"))
                                    {
                                        string valuelabeltb = tableTypeAndLabels[9].Value;
                                        foreach (var tablvalue in tableRows)
                                        {
                                            string rowKey2 = tablvalue.Key;
                                            List<KeyValuePair<string, string>> columnss = tablvalue.Value;
                                            foreach (var tablkey in columnss)
                                            {
                                                if (tablkey.Key == "Head working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "DCC Area")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "Working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "ID")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Name")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Position")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Division /Group / Section")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);
                                                                keylable = labelk.Key;

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }


                                            }
                                        }
                                        getadvance = ReplaceDataProcessBC(getadvance, listvalue, keylable, valuelabeltb);
                                        break;
                                    }
                                    if (column.Value == "11 : Production Support Center Group" || column.Value.Contains("11 :"))
                                    {
                                        string valuelabeltb = tableTypeAndLabels[10].Value;
                                        foreach (var tablvalue in tableRows)
                                        {
                                            string rowKey2 = tablvalue.Key;
                                            List<KeyValuePair<string, string>> columnss = tablvalue.Value;
                                            foreach (var tablkey in columnss)
                                            {
                                                if (tablkey.Key == "Head working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "DCC Area")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "Working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "ID")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Name")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Position")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Division /Group / Section")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);
                                                                keylable = labelk.Key;

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }


                                            }
                                        }
                                        getadvance = ReplaceDataProcessBC(getadvance, listvalue, keylable, valuelabeltb);
                                        break;
                                    }
                                    if (column.Value == "12 : Production Development Group" || column.Value.Contains("12 :"))
                                    {
                                        string valuelabeltb = tableTypeAndLabels[11].Value;
                                        foreach (var tablvalue in tableRows)
                                        {
                                            string rowKey2 = tablvalue.Key;
                                            List<KeyValuePair<string, string>> columnss = tablvalue.Value;
                                            foreach (var tablkey in columnss)
                                            {
                                                if (tablkey.Key == "Head working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "DCC Area")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "Working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "ID")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Name")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Position")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Division /Group / Section")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);
                                                                keylable = labelk.Key;

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }


                                            }
                                        }
                                        getadvance = ReplaceDataProcessBC(getadvance, listvalue, keylable, valuelabeltb);
                                        break;
                                    }
                                    if (column.Value == "13 : New Model Quality Assurance Group & Market Information Control Group & Quality Promotion & Planning Group & Technical Regulations Control Group" || column.Value.Contains("13 :"))
                                    {
                                        string valuelabeltb = tableTypeAndLabels[12].Value;
                                        foreach (var tablvalue in tableRows)
                                        {
                                            string rowKey2 = tablvalue.Key;
                                            List<KeyValuePair<string, string>> columnss = tablvalue.Value;
                                            foreach (var tablkey in columnss)
                                            {
                                                if (tablkey.Key == "Head working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "DCC Area")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "Working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "ID")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Name")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Position")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Division /Group / Section")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);
                                                                keylable = labelk.Key;

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }


                                            }
                                        }
                                        getadvance = ReplaceDataProcessBC(getadvance, listvalue, keylable, valuelabeltb);
                                        break;
                                    }
                                    if (column.Value == "14 : Product Quality Improvement Group" || column.Value.Contains("14 :"))
                                    {
                                        string valuelabeltb = tableTypeAndLabels[13].Value;
                                        foreach (var tablvalue in tableRows)
                                        {
                                            string rowKey2 = tablvalue.Key;
                                            List<KeyValuePair<string, string>> columnss = tablvalue.Value;
                                            foreach (var tablkey in columnss)
                                            {
                                                if (tablkey.Key == "Head working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "DCC Area")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "Working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "ID")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Name")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Position")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Division /Group / Section")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);
                                                                keylable = labelk.Key;

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }


                                            }
                                        }
                                        getadvance = ReplaceDataProcessBC(getadvance, listvalue, keylable, valuelabeltb);
                                        break;
                                    }
                                    if (column.Value == "15 : Production Planning & Control Division" || column.Value.Contains("15 :"))
                                    {
                                        string valuelabeltb = tableTypeAndLabels[14].Value;
                                        foreach (var tablvalue in tableRows)
                                        {
                                            string rowKey2 = tablvalue.Key;
                                            List<KeyValuePair<string, string>> columnss = tablvalue.Value;
                                            foreach (var tablkey in columnss)
                                            {
                                                if (tablkey.Key == "Head working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "DCC Area")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "Working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "ID")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Name")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Position")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Division /Group / Section")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);
                                                                keylable = labelk.Key;

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }


                                            }
                                        }
                                        getadvance = ReplaceDataProcessBC(getadvance, listvalue, keylable, valuelabeltb);
                                        break;
                                    }
                                    if (column.Value == "16 : Procurement Division" || column.Value.Contains("16 :"))
                                    {
                                        string valuelabeltb = tableTypeAndLabels[15].Value;
                                        foreach (var tablvalue in tableRows)
                                        {
                                            string rowKey2 = tablvalue.Key;
                                            List<KeyValuePair<string, string>> columnss = tablvalue.Value;
                                            foreach (var tablkey in columnss)
                                            {
                                                if (tablkey.Key == "Head working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "DCC Area")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "Working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "ID")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Name")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Position")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Division /Group / Section")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);
                                                                keylable = labelk.Key;

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }


                                            }
                                        }
                                        getadvance = ReplaceDataProcessBC(getadvance, listvalue, keylable, valuelabeltb);
                                        break;
                                    }
                                    if (column.Value == "17 : Assembly Engineering Division" || column.Value.Contains("17 :"))
                                    {
                                        string valuelabeltb = tableTypeAndLabels[16].Value;
                                        foreach (var tablvalue in tableRows)
                                        {
                                            string rowKey2 = tablvalue.Key;
                                            List<KeyValuePair<string, string>> columnss = tablvalue.Value;
                                            foreach (var tablkey in columnss)
                                            {
                                                if (tablkey.Key == "Head working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "DCC Area")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "Working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "ID")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Name")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Position")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Division /Group / Section")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);
                                                                keylable = labelk.Key;

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }


                                            }
                                        }
                                        getadvance = ReplaceDataProcessBC(getadvance, listvalue, keylable, valuelabeltb);
                                        break;
                                    }
                                    if (column.Value == "18 : Parts Engineering Group" || column.Value.Contains("18 :"))
                                    {
                                        string valuelabeltb = tableTypeAndLabels[17].Value;
                                        foreach (var tablvalue in tableRows)
                                        {
                                            string rowKey2 = tablvalue.Key;
                                            List<KeyValuePair<string, string>> columnss = tablvalue.Value;
                                            foreach (var tablkey in columnss)
                                            {
                                                if (tablkey.Key == "Head working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "DCC Area")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "Working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "ID")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Name")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Position")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Division /Group / Section")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);
                                                                keylable = labelk.Key;

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }


                                            }
                                        }
                                        getadvance = ReplaceDataProcessBC(getadvance, listvalue, keylable, valuelabeltb);
                                        break;
                                    }
                                    if (column.Value == "19 : General Affairs Division & Corporate Legal Group" || column.Value.Contains("19 :"))
                                    {
                                        string valuelabeltb = tableTypeAndLabels[18].Value;
                                        foreach (var tablvalue in tableRows)
                                        {
                                            string rowKey2 = tablvalue.Key;
                                            List<KeyValuePair<string, string>> columnss = tablvalue.Value;
                                            foreach (var tablkey in columnss)
                                            {
                                                if (tablkey.Key == "Head working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "DCC Area")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "Working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "ID")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Name")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Position")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Division /Group / Section")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);
                                                                keylable = labelk.Key;

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }


                                            }
                                        }
                                        getadvance = ReplaceDataProcessBC(getadvance, listvalue, keylable, valuelabeltb);
                                        break;
                                    }
                                    if (column.Value == "20 : Human Resources Division" || column.Value.Contains("20 :"))
                                    {
                                        string valuelabeltb = tableTypeAndLabels[19].Value;
                                        foreach (var tablvalue in tableRows)
                                        {
                                            string rowKey2 = tablvalue.Key;
                                            List<KeyValuePair<string, string>> columnss = tablvalue.Value;
                                            foreach (var tablkey in columnss)
                                            {
                                                if (tablkey.Key == "Head working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "DCC Area")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "Working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "ID")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Name")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Position")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Division /Group / Section")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);
                                                                keylable = labelk.Key;

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }


                                            }
                                        }
                                        getadvance = ReplaceDataProcessBC(getadvance, listvalue, keylable, valuelabeltb);
                                        break;
                                    }
                                    if (column.Value == "21 : Information Technology Division & Corporate Planning Division" || column.Value.Contains("21 :"))
                                    {
                                        string valuelabeltb = tableTypeAndLabels[20].Value;
                                        foreach (var tablvalue in tableRows)
                                        {
                                            string rowKey2 = tablvalue.Key;
                                            List<KeyValuePair<string, string>> columnss = tablvalue.Value;
                                            foreach (var tablkey in columnss)
                                            {
                                                if (tablkey.Key == "Head working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "DCC Area")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "Working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "ID")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Name")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Position")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Division /Group / Section")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);
                                                                keylable = labelk.Key;

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }


                                            }
                                        }
                                        getadvance = ReplaceDataProcessBC(getadvance, listvalue, keylable, valuelabeltb);
                                        break;
                                    }
                                    if (column.Value == "22 : Finance & Credit Management Division &IA Group" || column.Value.Contains("22 :"))
                                    {
                                        string valuelabeltb = tableTypeAndLabels[21].Value;
                                        foreach (var tablvalue in tableRows)
                                        {
                                            string rowKey2 = tablvalue.Key;
                                            List<KeyValuePair<string, string>> columnss = tablvalue.Value;
                                            foreach (var tablkey in columnss)
                                            {
                                                if (tablkey.Key == "Head working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "DCC Area")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "Working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "ID")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Name")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Position")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Division /Group / Section")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);
                                                                keylable = labelk.Key;

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }


                                            }
                                        }
                                        getadvance = ReplaceDataProcessBC(getadvance, listvalue, keylable, valuelabeltb);
                                        break;
                                    }
                                    if (column.Value == "23 : Accounting & Financial Planning Division" || column.Value.Contains("23 :"))
                                    {
                                        string valuelabeltb = tableTypeAndLabels[22].Value;
                                        foreach (var tablvalue in tableRows)
                                        {
                                            string rowKey2 = tablvalue.Key;
                                            List<KeyValuePair<string, string>> columnss = tablvalue.Value;
                                            foreach (var tablkey in columnss)
                                            {
                                                if (tablkey.Key == "Head working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "DCC Area")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "Working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "ID")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Name")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Position")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Division /Group / Section")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);
                                                                keylable = labelk.Key;

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }


                                            }
                                        }
                                        getadvance = ReplaceDataProcessBC(getadvance, listvalue, keylable, valuelabeltb);
                                        break;
                                    }
                                    if (column.Value == "24 : Sales Administration Group & SCM Planning Group & Fleet & Finance Division" || column.Value.Contains("24 :") || column.Value.Contains("24:"))
                                    {
                                        string valuelabeltb = tableTypeAndLabels[23].Value;
                                        foreach (var tablvalue in tableRows)
                                        {
                                            string rowKey2 = tablvalue.Key;
                                            List<KeyValuePair<string, string>> columnss = tablvalue.Value;
                                            foreach (var tablkey in columnss)
                                            {
                                                if (tablkey.Key == "Head working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "DCC Area")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "Working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "ID")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Name")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Position")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Division /Group / Section")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);
                                                                keylable = labelk.Key;

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }


                                            }
                                        }
                                        getadvance = ReplaceDataProcessBC(getadvance, listvalue, keylable, valuelabeltb);
                                        break;
                                    }
                                    if (column.Value == "25 : Merchandise Operation & Engineering Group & Big Bike Business Division & Motor Sport Planning & activities Group & Racing Group & Customer Satisfaction Group" || column.Value.Contains("25 :"))
                                    {
                                        string valuelabeltb = tableTypeAndLabels[24].Value;
                                        foreach (var tablvalue in tableRows)
                                        {
                                            string rowKey2 = tablvalue.Key;
                                            List<KeyValuePair<string, string>> columnss = tablvalue.Value;
                                            foreach (var tablkey in columnss)
                                            {
                                                if (tablkey.Key == "Head working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "DCC Area")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "Working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "ID")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Name")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Position")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Division /Group / Section")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);
                                                                keylable = labelk.Key;

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }


                                            }
                                        }
                                        getadvance = ReplaceDataProcessBC(getadvance, listvalue, keylable, valuelabeltb);
                                        break;
                                    }
                                    if (column.Value == "26 : Service Division" || column.Value.Contains("26 :"))
                                    {
                                        string valuelabeltb = tableTypeAndLabels[25].Value;
                                        foreach (var tablvalue in tableRows)
                                        {
                                            string rowKey2 = tablvalue.Key;
                                            List<KeyValuePair<string, string>> columnss = tablvalue.Value;
                                            foreach (var tablkey in columnss)
                                            {
                                                if (tablkey.Key == "Head working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "DCC Area")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "Working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "ID")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Name")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Position")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Division /Group / Section")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);
                                                                keylable = labelk.Key;

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }


                                            }
                                        }
                                        getadvance = ReplaceDataProcessBC(getadvance, listvalue, keylable, valuelabeltb);
                                        break;
                                    }
                                    if (column.Value == "27 : Spare Parts Division" || column.Value.Contains("27 :"))
                                    {
                                        string valuelabeltb = tableTypeAndLabels[26].Value;
                                        foreach (var tablvalue in tableRows)
                                        {
                                            string rowKey2 = tablvalue.Key;
                                            List<KeyValuePair<string, string>> columnss = tablvalue.Value;
                                            foreach (var tablkey in columnss)
                                            {
                                                if (tablkey.Key == "Head working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "DCC Area")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "Working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "ID")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Name")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Position")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Division /Group / Section")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);
                                                                keylable = labelk.Key;

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }


                                            }
                                        }
                                        getadvance = ReplaceDataProcessBC(getadvance, listvalue, keylable, valuelabeltb);
                                        break;
                                    }
                                    if (column.Value == "28 : Marketing Planning & Product Planning Division & Marine & Golf Car Group" || column.Value.Contains("28 :"))
                                    {
                                        string valuelabeltb = tableTypeAndLabels[27].Value;
                                        foreach (var tablvalue in tableRows)
                                        {
                                            string rowKey2 = tablvalue.Key;
                                            List<KeyValuePair<string, string>> columnss = tablvalue.Value;
                                            foreach (var tablkey in columnss)
                                            {
                                                if (tablkey.Key == "Head working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "DCC Area")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "Working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "ID")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Name")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Position")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Division /Group / Section")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);
                                                                keylable = labelk.Key;

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }


                                            }
                                        }
                                        getadvance = ReplaceDataProcessBC(getadvance, listvalue, keylable, valuelabeltb);
                                        break;
                                    }
                                    if (column.Value == "29 : Marketing-AT& Marketing support Division & Marketing Sports & Moped Division" || column.Value.Contains("29 :"))
                                    {
                                        string valuelabeltb = tableTypeAndLabels[28].Value;
                                        foreach (var tablvalue in tableRows)
                                        {
                                            string rowKey2 = tablvalue.Key;
                                            List<KeyValuePair<string, string>> columnss = tablvalue.Value;
                                            foreach (var tablkey in columnss)
                                            {
                                                if (tablkey.Key == "Head working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "DCC Area")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "Working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "ID")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Name")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Position")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Division /Group / Section")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);
                                                                keylable = labelk.Key;

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }


                                            }
                                        }
                                        getadvance = ReplaceDataProcessBC(getadvance, listvalue, keylable, valuelabeltb);
                                        break;
                                    }
                                    if (column.Value == "30 : Motor Sport Division" || column.Value.Contains("30 :"))
                                    {
                                        string valuelabeltb = tableTypeAndLabels[29].Value;
                                        foreach (var tablvalue in tableRows)
                                        {
                                            string rowKey2 = tablvalue.Key;
                                            List<KeyValuePair<string, string>> columnss = tablvalue.Value;
                                            foreach (var tablkey in columnss)
                                            {
                                                if (tablkey.Key == "Head working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "DCC Area")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "Working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "ID")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Name")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Position")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Division /Group / Section")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);
                                                                keylable = labelk.Key;

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }


                                            }
                                        }
                                        getadvance = ReplaceDataProcessBC(getadvance, listvalue, keylable, valuelabeltb);
                                        break;
                                    }
                                    if (column.Value == "31 : Low Cost Automation & AI Group" || column.Value.Contains("31 :"))
                                    {
                                        string valuelabeltb = tableTypeAndLabels[30].Value;
                                        foreach (var tablvalue in tableRows)
                                        {
                                            string rowKey2 = tablvalue.Key;
                                            List<KeyValuePair<string, string>> columnss = tablvalue.Value;
                                            foreach (var tablkey in columnss)
                                            {
                                                if (tablkey.Key == "Head working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "DCC Area")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "Working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "ID")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Name")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Position")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Division /Group / Section")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);
                                                                keylable = labelk.Key;

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }


                                            }
                                        }
                                        getadvance = ReplaceDataProcessBC(getadvance, listvalue, keylable, valuelabeltb);
                                        break;
                                    }
                                    if (column.Value == "32 : Logistics Innovation Division" || column.Value.Contains("32 :"))
                                    {
                                        string valuelabeltb = tableTypeAndLabels[31].Value;
                                        foreach (var tablvalue in tableRows)
                                        {
                                            string rowKey2 = tablvalue.Key;
                                            List<KeyValuePair<string, string>> columnss = tablvalue.Value;
                                            foreach (var tablkey in columnss)
                                            {
                                                if (tablkey.Key == "Head working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "DCC Area")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "Working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "ID")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Name")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Position")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Division /Group / Section")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);
                                                                keylable = labelk.Key;

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }


                                            }
                                        }
                                        getadvance = ReplaceDataProcessBC(getadvance, listvalue, keylable, valuelabeltb);
                                        break;
                                    }
                                    if (column.Value == "33 : YRA Division" || column.Value.Contains("33 :"))
                                    {
                                        string valuelabeltb = tableTypeAndLabels[32].Value;
                                        foreach (var tablvalue in tableRows)
                                        {
                                            string rowKey2 = tablvalue.Key;
                                            List<KeyValuePair<string, string>> columnss = tablvalue.Value;
                                            foreach (var tablkey in columnss)
                                            {
                                                if (tablkey.Key == "Head working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "DCC Area")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "Working team")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    }
                                                }
                                                if (tablkey.Key == "ID")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Name")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Position")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }
                                                if (tablkey.Key == "Division /Group / Section")
                                                {
                                                    foreach (var labelv in tableData)
                                                    {
                                                        foreach (var labelk in labelv)
                                                        {
                                                            if (labelk.Key == "Responsibilities")
                                                            {

                                                                listvalue.Add(tablkey.Value);
                                                                keylable = labelk.Key;

                                                                break;
                                                            }

                                                        }
                                                        break;
                                                    }
                                                }


                                            }
                                        }
                                        getadvance = ReplaceDataProcessBC(getadvance, listvalue, keylable, valuelabeltb);
                                        break;
                                    }
                                }
                                break;
                            }
                        }

                    }
                    var json = JsonConvert.DeserializeObject<dynamic>(getadvance);

                    string MAdvance = JsonConvert.SerializeObject(json, Formatting.None);
                    var EmpCurrent = dbWolf.ViewEmployees.Where(x => x.EmployeeCode == _EmpCode).FirstOrDefault();
                    var CurrentCom = dbWolf.MSTCompanies.Where(a => a.CompanyCode == EmpCurrent.CompanyCode).FirstOrDefault();

                    TRNMemo objMemo = new TRNMemo();
                    objMemo.StatusName = "Draft";
                    objMemo.CreatedDate = DateTime.Now;
                    objMemo.CreatedBy = EmpCurrent.NameEn;
                    objMemo.CreatorId = EmpCurrent.EmployeeId;
                    objMemo.RequesterId = EmpCurrent.EmployeeId;
                    objMemo.CNameTh = EmpCurrent.NameTh;
                    objMemo.CNameEn = EmpCurrent.NameEn;
                    objMemo.CPositionId = EmpCurrent.PositionId;
                    objMemo.CPositionTh = EmpCurrent.PositionNameTh;
                    objMemo.CPositionEn = EmpCurrent.PositionNameEn;
                    objMemo.CDepartmentId = EmpCurrent.DepartmentId;
                    objMemo.CDepartmentTh = EmpCurrent.DepartmentNameTh;
                    objMemo.CDepartmentEn = EmpCurrent.DepartmentNameEn;
                    objMemo.RNameTh = EmpCurrent.NameTh;
                    objMemo.RNameEn = EmpCurrent.NameEn;
                    objMemo.RPositionId = EmpCurrent.PositionId;
                    objMemo.RPositionTh = EmpCurrent.PositionNameTh;
                    objMemo.RPositionEn = EmpCurrent.PositionNameEn;
                    objMemo.RDepartmentId = EmpCurrent.DepartmentId;
                    objMemo.RDepartmentTh = EmpCurrent.DepartmentNameTh;
                    objMemo.RDepartmentEn = EmpCurrent.DepartmentNameEn;
                    objMemo.ModifiedDate = DateTime.Now;
                    objMemo.ModifiedBy = objMemo.ModifiedBy;
                    objMemo.TemplateId = Gettemid.TemplateId;
                    objMemo.TemplateName = Gettemid.TemplateName;
                    objMemo.GroupTemplateName = Gettemid.GroupTemplateName;
                    objMemo.RequestDate = DateTime.Now;
                    objMemo.PersonWaiting = EmpCurrent.NameEn;
                    objMemo.PersonWaitingId = EmpCurrent.EmployeeId;
                    objMemo.CompanyId = CurrentCom.CompanyId;
                    objMemo.CompanyName = CurrentCom.NameTh;
                    objMemo.MemoSubject = Gettemid.TemplateSubject;
                    objMemo.MAdvancveForm = MAdvance;
                    objMemo.TAdvanceForm = Gettemid.AdvanceForm;
                    objMemo.TemplateSubject = Gettemid.TemplateSubject;
                    objMemo.TemplateDetail = Guid.NewGuid().ToString().Replace("-", "");
                    objMemo.ToPerson = Gettemid.ToId;
                    objMemo.CcPerson = Gettemid.CcId;
                    objMemo.CurrentApprovalLevel = null;
                    objMemo.ProjectID = 0;
                    objMemo.Amount = 0;
                    objMemo.DocumentCode = GenControlRunning(EmpCurrent, Gettemid.DocumentCode, objMemo, dbWolf);
                    objMemo.DocumentNo = objMemo.DocumentCode;

                    dbWolf.TRNMemos.InsertOnSubmit(objMemo);
                    dbWolf.SubmitChanges();

                    List<ApprovalDetail> lineapprove = new List<ApprovalDetail>();

                    var gettemlineid = dbWolf.MSTTemLineApproves.Where(x => x.TemplateId == Gettemid.TemplateId).ToList();
                    int i = 1;
                    List<ViewEmployee> groubline = new List<ViewEmployee>();
                    foreach (var item in gettemlineid)
                    {
                        var gettemspecific = dbWolf.MSTTemSpecificApprovers.Where(x => x.TemLineId == item.TemLineId).ToList();
                        foreach (var getline in gettemspecific)
                        {
                            var rold = dbWolf.MSTUserPermissions.Where(x => x.RoleId == getline.EmployeeId).ToList();
                            foreach (var getemp in rold)
                            {
                                var emp = dbWolf.ViewEmployees.Where(x => x.EmployeeId == getemp.EmployeeId).FirstOrDefault();
                                TRNLineApprove trnLine = new TRNLineApprove();

                                trnLine.MemoId = objMemo.MemoId;
                                trnLine.Seq = i;
                                trnLine.EmployeeId = emp.EmployeeId;
                                trnLine.EmployeeCode = emp.EmployeeCode;
                                trnLine.NameTh = emp.NameTh;
                                trnLine.NameEn = emp.NameEn;
                                trnLine.PositionTH = emp.PositionNameTh;
                                trnLine.PositionEN = emp.PositionNameEn;
                                trnLine.SignatureId = 2019;
                                trnLine.SignatureTh = "อนุมัติ";
                                trnLine.SignatureEn = "Approved";
                                trnLine.IsActive = emp.IsActive;

                                dbWolf.TRNLineApproves.InsertOnSubmit(trnLine);
                                i++;
                            }

                        }
                        dbWolf.SubmitChanges();

                    }
                }

            }
            catch (Exception ex)
            {
                WriteLogFile(ex.Message);
                WriteLogFile(ex.StackTrace);
                WriteLogFile("Mempid catch :" + getmemoid.MemoId);
            }
        }
        public static TRNControlRunning InsertDocumentNo(DbWolfDataContext dbWolf)
        {
            var lastcontronlrunning = dbWolf.TRNControlRunnings.Where(z => z.TemplateId == 101 && z.Prefix.Length == 5).OrderBy(g => g.Running).LastOrDefault();

            string[] parts = lastcontronlrunning.RunningNumber.Split('-');
            int numberPart = int.Parse(parts[1]);
            numberPart += 1;
            string running = $"{DateTime.Now.Year.ToString()}-{numberPart:D3}";
            TRNControlRunning objControlRunning = new TRNControlRunning();
            objControlRunning.TemplateId = 101;
            objControlRunning.Digit = 3;
            objControlRunning.CreateBy = "1";
            objControlRunning.CreateDate = DateTime.Now;
            objControlRunning.RunningNumber = running;
            string number = "";
            string lastNumber = "";
            lastNumber = running.Split('-').Last();
            number = running.Replace($"{lastNumber}", "");
            objControlRunning.Prefix = number;
            string running1 = "";
            if (Regex.IsMatch(lastNumber, @"^(0?[1-9]0|[1-9]00)$"))
            {
                running1 = lastNumber.TrimStart('0');
            }
            else
            {
                running1 = lastNumber.TrimStart('0');
            }
            objControlRunning.Running = Convert.ToInt32(running1);
            dbWolf.TRNControlRunnings.InsertOnSubmit(objControlRunning);
            dbWolf.SubmitChanges();
            return objControlRunning;
        }
        public static List<ApprovalDetail> GetLineapprove(List<ViewEmployee> lstEmp, int TemplateID, List<string> Type, string heardlabel)
        {
            //Type = Type.Replace("\\", "");

            Console.WriteLine("Type : " + Type);

            DbWolfDataContext dbWolf = new DbWolfDataContext(dbConnectionStringWolf);
            List<MSTTemplateLogic> LogicID = new List<MSTTemplateLogic>();
            string logic = string.Empty;
            if (dbWolf.Connection.State == ConnectionState.Open)
            {
                dbWolf.Connection.Close();
                dbWolf.Connection.Open();
            }
            else
            {

                dbWolf.Connection.Open();
                LogicID = dbWolf.MSTTemplateLogics.Where(a => a.TemplateId == TemplateID & a.logictype == "datalineapprove").ToList();
                foreach (var loadlocgic in LogicID)
                {
                    var logicType = JObject.Parse(loadlocgic.jsonvalue);

                    string jlabel = logicType["label"].ToString();
                    if (heardlabel.Contains(jlabel))
                    {
                        logic = loadlocgic.logicid.ToString();
                        break;
                    }
                }

            }
            string jsonFormatCondition = "{'logicid':'" + logic + "','conditions':[";

            foreach (var item in Type)
            {

                var parts = item.Split('|');
                if (parts.Length >= 2)
                {
                    jsonFormatCondition += "{'label':'" + parts[0].Replace("'", "\\'") + "','value':'" + parts[1].Replace("'", "\\'") + "'},";
                }
                else
                {
                    jsonFormatCondition += "{'label':'" + heardlabel.Replace("'", "\\'") + "','value':'" + item.Replace("'", "\\'") + "'},";
                }



            }

            // Remove the trailing comma if not needed
            if (jsonFormatCondition.EndsWith(","))
            {
                jsonFormatCondition = jsonFormatCondition.Remove(jsonFormatCondition.Length - 1);
            }

            jsonFormatCondition += "]}";

            // Replace single quotes with double quotes to make it a valid JSON string
            jsonFormatCondition = jsonFormatCondition.Replace("'", "\"");
            string FormatCondition = jsonFormatCondition;


            Console.WriteLine("FormatCondition : " + FormatCondition);

            ////return lstapprovalDetails;
            TemplateDetailFormPage templateDetailFormPage = new TemplateDetailFormPage();

            try
            {

                templateDetailFormPage.connectionString = dbConnectionStringWolf;
                templateDetailFormPage.lstTRNLineApprove = new List<ApprovalDetail>();
                templateDetailFormPage.templateForm = JsonConvert.DeserializeObject<CustomTemplate>(postAPI($"api/Template/TemplateByID", new CustomTemplate() { connectionString = dbConnectionStringWolf, TemplateId = TemplateID }));
                templateDetailFormPage.VEmployee = ConvertEmpToCustom(lstEmp.First());
                templateDetailFormPage.JsonCondition = FormatCondition;
                Console.WriteLine("Get TemplateLogic ID : " + TemplateID);

            }
            catch (Exception ex)
            {
                Console.WriteLine("GetLineapprove error :" + ex.Message.ToString());
                throw;
            }

            WriteLogFile(templateDetailFormPage.ToJson());
            return JsonConvert.DeserializeObject<List<ApprovalDetail>>(postAPI($"api/LineApprove/LineApproveWithTemplate", templateDetailFormPage));

        }
        public static string ReplaceDataProcessBC(string DestAdvanceForm, List<string> Value, string label, string valuelabeltb)
        {
            JObject jsonAdvanceForm = JObject.Parse(DestAdvanceForm);
            JArray itemsArray = (JArray)jsonAdvanceForm["items"];
            if (itemsArray == null) return DestAdvanceForm;

            foreach (JObject jItems in itemsArray)
            {
                JArray jLayoutArray = (JArray)jItems["layout"];
                if (jLayoutArray == null) continue;

                foreach (JObject jLayout in jLayoutArray)
                {
                    JObject jTemplate = (JObject)jLayout["template"];
                    if (jTemplate == null) continue;

                    string currentType = (string)jTemplate["type"];
                    string currentLabel = (string)jTemplate["label"];

                    if (currentType == "tb" && currentLabel == valuelabeltb)
                    {
                        JArray newRows = new JArray();
                        for (int i = 0; i < Value.Count; i += 5)
                        {
                            JArray row = new JArray();
                            for (int j = 0; j < 5; j++)
                            {
                                JObject cell = new JObject();
                                cell["value"] = i + j < Value.Count ? Value[i + j] : null;
                                row.Add(cell);
                            }
                            newRows.Add(row);
                        }

                        JObject dataObject = new JObject();
                        dataObject["row"] = newRows;
                        jLayout["data"] = dataObject;
                    }
                }
            }

            return jsonAdvanceForm.ToString();
        }
        public static string GenControlRunning(ViewEmployee Emp, string DocumentCode, TRNMemo objTRNMemo, DbWolfDataContext db)
        {
            string TempCode = DocumentCode;
            String sPrefixDocNo = $"{TempCode}-{DateTime.Now.Year.ToString()}-";
            int iRunning = 1;
            List<TRNMemo> temp = db.TRNMemos.Where(a => a.DocumentNo.ToUpper().Contains(sPrefixDocNo.ToUpper())).ToList();
            if (temp.Count > 0)
            {
                String sLastDocumentNo = temp.OrderBy(a => a.DocumentNo).Last().DocumentNo;
                if (!String.IsNullOrEmpty(sLastDocumentNo))
                {
                    List<String> list_LastDocumentNo = sLastDocumentNo.Split('-').ToList();

                    if (list_LastDocumentNo.Count >= 3)
                    {
                        iRunning = checkDataIntIsNull(list_LastDocumentNo[list_LastDocumentNo.Count - 1]) + 1;
                    }
                }
            }
            String sDocumentNo = $"{sPrefixDocNo}{iRunning.ToString().PadLeft(3, '0')}";
            TRNControlRunning controrun = new TRNControlRunning();
            controrun.TemplateId = objTRNMemo.TemplateId;
            controrun.Prefix = TempCode + "-";
            controrun.Digit = 3;
            controrun.Running = iRunning;
            controrun.CreateBy = "1";
            controrun.CreateDate = DateTime.Now;
            controrun.RunningNumber = sDocumentNo;

            db.TRNControlRunnings.InsertOnSubmit(controrun);
            db.SubmitChanges();

            try
            {

                var mstMasterDataList = db.MSTMasterDatas.Where(a => a.MasterType == "DocNo").ToList();

                if (mstMasterDataList != null)
                    if (mstMasterDataList.Count() > 0)
                    {
                        var getCompany = db.MSTCompanies.Where(a => a.CompanyId == objTRNMemo.CompanyId).ToList();
                        var getDepartment = db.MSTDepartments.Where(a => a.DepartmentId == Emp.DepartmentId).ToList();
                        var getDivision = db.MSTDivisions.Where(a => a.DivisionId == Emp.DivisionId).ToList();

                        string CompanyCode = "";
                        string DepartmentCode = "";
                        string DivisionCode = "";
                        if (getCompany != null)
                            if (!string.IsNullOrWhiteSpace(getCompany.First().CompanyCode)) CompanyCode = getCompany.First().CompanyCode;
                        if (DepartmentCode != null)
                            if (!string.IsNullOrWhiteSpace(getDepartment.First().DepartmentCode)) DepartmentCode = getDepartment.First().DepartmentCode;
                        if (DivisionCode != null)
                        {
                            if (getDivision.Count > 0)
                                if (!string.IsNullOrWhiteSpace(getDivision.First().DivisionCode)) DivisionCode = getDivision.First().DivisionCode;
                        }
                        foreach (var getMaster in mstMasterDataList)
                        {
                            if (!string.IsNullOrWhiteSpace(getMaster.Value2))
                            {
                                var Tid_array = getMaster.Value2.Split('|');
                                string FixDoc = getMaster.Value1;
                                if (Tid_array.Count() > 0)
                                {
                                    if (Tid_array.Contains(objTRNMemo.TemplateId.ToString()))
                                    {
                                        sDocumentNo = DocNoGenerate(FixDoc, TempCode, CompanyCode, DepartmentCode, DivisionCode, db);
                                    }
                                }
                            }
                            else
                            {
                                string FixDoc = getMaster.Value1;
                                sDocumentNo = DocNoGenerate(FixDoc, TempCode, CompanyCode, DepartmentCode, DivisionCode, db);
                            }
                        }
                    }
            }
            catch (Exception ex) { }

            return sDocumentNo;
        }
        public static string DocNoGenerate(string FixDoc, string DocCode, string CCode, string DCode, string DSCode, DbWolfDataContext db)
        {
            string sDocumentNo = "";
            int iRunning;
            if (!string.IsNullOrWhiteSpace(FixDoc))
            {
                string y4 = DateTime.Now.ToString("yyyy");
                string y2 = DateTime.Now.ToString("yy");
                string CompanyCode = CCode;
                string DepartmentCode = DCode;
                string DivisionCode = DSCode;
                string FixCode = FixDoc;
                FixCode = FixCode.Replace("[CompanyCode]", CompanyCode);
                FixCode = FixCode.Replace("[DepartmentCode]", DepartmentCode);
                FixCode = FixCode.Replace("[DocumentCode]", DocCode);
                FixCode = FixCode.Replace("[DivisionCode]", DivisionCode);

                FixCode = FixCode.Replace("[YYYY]", y4);
                FixCode = FixCode.Replace("[YY]", y2);
                sDocumentNo = FixCode;
                List<TRNMemo> tempfixDoc = db.TRNMemos.Where(a => a.DocumentNo.ToUpper().Contains(sDocumentNo.ToUpper())).ToList();


                List<TRNMemo> tempfixDocByYear = db.TRNMemos.ToList();

                tempfixDocByYear = tempfixDocByYear.FindAll(a => a.DocumentNo != ("Auto Generate") & Convert.ToDateTime(a.RequestDate).Year.ToString().Equals(y4)).ToList();

                if (tempfixDocByYear.Count > 0)
                {
                    tempfixDocByYear = tempfixDocByYear.OrderByDescending(a => a.MemoId).ToList();

                    String sLastDocumentNofix = tempfixDocByYear.First().DocumentNo;
                    if (!String.IsNullOrEmpty(sLastDocumentNofix))
                    {
                        List<String> list_LastDocumentNofix = sLastDocumentNofix.Split('-').ToList();

                        if (list_LastDocumentNofix.Count >= 3)
                        {
                            iRunning = checkDataIntIsNull(list_LastDocumentNofix[list_LastDocumentNofix.Count - 1]) + 1;
                            sDocumentNo = $"{sDocumentNo}-{iRunning.ToString().PadLeft(3, '0')}";




                        }
                    }
                }
                else
                {
                    sDocumentNo = $"{sDocumentNo}-{1.ToString().PadLeft(3, '0')}";

                }
            }
            return sDocumentNo;
        }
        public static string PushValue(string DestAdvanceForm, string Value, string label)
        {
            JObject jsonAdvanceForm = JObject.Parse(DestAdvanceForm);
            JArray itemsArray = (JArray)jsonAdvanceForm["items"];
            foreach (JObject jItems in itemsArray)
            {
                JArray jLayoutArray = (JArray)jItems["layout"];

                foreach (JObject jLayout in jLayoutArray)
                {
                    // Existing code to handle the first and second template
                    JObject jTemplate = (JObject)jLayout["template"];
                    if (jTemplate != null && (string)jTemplate["label"] == label)
                    {
                        JObject jData = (JObject)jLayout["data"];
                        if (jData != null)
                        {
                            jData["value"] = Value;
                        }
                    }

                    // Additional code to handle nested columns in 'attribute'
                    JObject jAttribute = (JObject)jTemplate["attribute"];
                    if (jAttribute != null)
                    {
                        JArray columns = (JArray)jAttribute["column"];
                        if (columns != null && columns.Count > 0)
                        {
                            foreach (JObject column in columns)
                            {
                                if ((string)column["label"] == label)
                                {
                                    JObject control = (JObject)column["control"];
                                    JObject jData = (JObject)control["data"];
                                    if (jData != null)
                                    {
                                        jData["value"] = Value;
                                    }
                                }
                            }
                        }

                    }
                }
            }
            return JsonConvert.SerializeObject(jsonAdvanceForm);
        }
        public static CustomViewEmployee ConvertEmpToCustom(ViewEmployee viewEmployee)
        {
            CustomViewEmployee oResult = new CustomViewEmployee();
            oResult.EmployeeId = viewEmployee.EmployeeId;
            oResult.EmployeeCode = viewEmployee.EmployeeCode;
            oResult.Username = viewEmployee.Username;
            oResult.NameTh = viewEmployee.NameTh;
            oResult.NameEn = viewEmployee.NameEn;
            oResult.Email = viewEmployee.Email;
            oResult.IsActive = viewEmployee.IsActive;
            oResult.PositionId = viewEmployee.PositionId;
            oResult.PositionNameTh = viewEmployee.PositionNameTh;
            oResult.PositionNameEn = viewEmployee.PositionNameEn;
            oResult.DepartmentId = viewEmployee.DepartmentId;
            oResult.DepartmentNameTh = viewEmployee.DepartmentNameTh;
            oResult.DepartmentNameEn = viewEmployee.DepartmentNameEn;
            oResult.SignPicPath = viewEmployee.SignPicPath;
            oResult.Lang = viewEmployee.Lang;
            //AccountId = viewEmployee.AccountId;
            oResult.AccountCode = viewEmployee.AccountCode;
            oResult.AccountName = viewEmployee.AccountName;
            oResult.DefaultLang = viewEmployee.DefaultLang;
            oResult.RegisteredDate = DateTimeHelper.DateTimeToCustomClassString(viewEmployee.RegisteredDate);
            oResult.ExpiredDate = DateTimeHelper.DateTimeToCustomClassString(viewEmployee.ExpiredDate);
            oResult.CreatedDate = DateTimeHelper.DateTimeToCustomClassString(viewEmployee.CreatedDate);
            oResult.CreatedBy = viewEmployee.CreatedBy;
            oResult.ModifiedDate = DateTimeHelper.DateTimeToCustomClassString(viewEmployee.ModifiedDate);
            oResult.ModifiedBy = viewEmployee.ModifiedBy;
            oResult.ReportToEmpCode = viewEmployee.ReportToEmpCode;
            oResult.DivisionId = viewEmployee.DivisionId;
            oResult.DivisionNameTh = viewEmployee.DivisionNameTh;
            oResult.DivisionNameEn = viewEmployee.DivisionNameEn;
            oResult.ADTitle = viewEmployee.ADTitle;

            return oResult;
        }
        public static int checkDataIntIsNull(object Input)
        {
            int Results = 0;
            if (Input != null)
                int.TryParse(Input.ToString().Replace(",", ""), out Results);

            return Results;
        }
        private static string _LogFile
        {
            get
            {
                var LogFile = System.Configuration.ConfigurationSettings.AppSettings["LogFile"];
                if (!string.IsNullOrEmpty(LogFile))
                {
                    return (LogFile);
                }
                return string.Empty;
            }
        }
        private static string DataTime_Run
        {
            get
            {
                var DataTime_Run = System.Configuration.ConfigurationSettings.AppSettings["DataTime_Run"];
                if (!string.IsNullOrEmpty(DataTime_Run))
                {
                    return (DataTime_Run);
                }
                return string.Empty;
            }
        }
        private static string _BaseAPI
        {
            get
            {
                var BaseAPI = System.Configuration.ConfigurationSettings.AppSettings["BaseAPI"];
                if (!string.IsNullOrEmpty(BaseAPI))
                {
                    return (BaseAPI);
                }
                return string.Empty;
            }
        }
        private static int iIntervalTime
        {
            //ตั้งค่าเวลา
            get
            {
                var IntervalTime = System.Configuration.ConfigurationSettings.AppSettings["IntervalTimeMinute"];
                if (!string.IsNullOrEmpty(IntervalTime))
                {
                    return Convert.ToInt32(IntervalTime);
                }
                return -10;
            }
        }
        private static string _EmpCode
        {
            get
            {
                var EmpCode = System.Configuration.ConfigurationSettings.AppSettings["EmpCode"];
                if (!string.IsNullOrEmpty(EmpCode))
                {
                    return (EmpCode);
                }
                return string.Empty;
            }
        }
        public static List<CustomViewEmployee> GetEmployeeByEmpCode(CustomViewEmployee empCode)
        {
            List<CustomViewEmployee> obj = JsonConvert.DeserializeObject<List<CustomViewEmployee>>(postAPI($"api/Employee/EmployeeByEmpCode", empCode));
            return obj == null ? new List<CustomViewEmployee>() : obj;
        }
        public static void WriteLogFile(String iText)
        {

            String LogFilePath = String.Format("{0}{1}_OrderLog.txt", _LogFile, DateTime.Now.ToString("yyyyMMdd"));

            try
            {
                using (System.IO.StreamWriter outfile = new System.IO.StreamWriter(LogFilePath, true))
                {
                    System.Text.StringBuilder sbLog = new System.Text.StringBuilder();

                    String[] ListText = iText.Split('|').ToArray();

                    foreach (String s in ListText)
                    {
                        sbLog.AppendLine(s);
                    }

                    outfile.WriteLine(sbLog.ToString());
                }
            }
            catch { }
        }
        public static string postAPI(string subUri, Object obj)
        {
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    client.BaseAddress = new Uri(_BaseAPI);
                    client.DefaultRequestHeaders.Accept.Clear();
                    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                    string json = new JavaScriptSerializer().Serialize(obj);
                    StringContent content = new StringContent(json, Encoding.UTF8, "application/json");

                    Task<HttpResponseMessage> response = client.PostAsync(subUri, content);

                    if (response.Result.IsSuccessStatusCode)
                    {
                        return response.Result.Content.ReadAsStringAsync().Result;
                    }
                    else
                    {
                        return "Not Found";
                    }
                }
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }
    }
}
