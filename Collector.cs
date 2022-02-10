using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.IO;
using System.Threading;
using System.Resources;
using System.Collections;
using Google.Apis.Drive.v3.Data;
using System.Web;
using Google.Apis.Drive.v3;
using System.Net;
using System.Globalization;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;

namespace DataCollector
{
    public partial class DataCollector : Form
    {

        static string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
        static string ApplicationName = "Get Google SheetData with Google Sheets API";
        public DataCollector()
        {
            InitializeComponent();
        }

        private void DataCollector_Load(object sender, EventArgs e)
        {
        }

        void Load_Data(string id, string cases_array, DataGridView grid, System.Windows.Forms.Button button)
        {
            grid.Rows.Clear();

            UserCredential credential;
            SheetsService service = null;
            try
            {
                using (var stream =
                    new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
                {
                    string credPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal);
                    credPath = Path.Combine(credPath, ".credentials/sheets.googleapis.com-dotnet-quickstart.json");
                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.FromStream(stream).Secrets,
                        Scopes,
                        "user",
                        CancellationToken.None,
                        new FileDataStore(credPath, true)).Result;
                    Console.WriteLine("Credential file saved to: " + credPath);
                }

                if (credential == null)
                {
                    this.Close();
                }

                // Create Google Sheets API service.
                service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });
            } catch(Exception ex)
            {
                MessageBox.Show("資料傳輸出現問題!", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            String[] cases_list = cases_array.Split('/');
            IList<IList<Object>> cases = null;
            IList<IList<Object>> cases2 = null;
            IList<IList<Object>> time = null;
            String spreadsheetId = id;
            String range = "表單回應 1!H2:H50000";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId, range);
            cases = request.Execute().Values;

            range = "表單回應 1!C2:C50000";
            request = service.Spreadsheets.Values.Get(spreadsheetId, range);
            cases2 = request.Execute().Values;

            range = "表單回應 1!A2:A50000";
            request = service.Spreadsheets.Values.Get(spreadsheetId, range);
            time = request.Execute().Values;

            if (cases != null && cases.Count > 0 || cases2 != null && cases2.Count > 0)
            {
                List<string> unresponded = new List<string>();
                unresponded.AddRange(cases_list);
                Dictionary<string, DateTime> responded = new Dictionary<string, DateTime>();

                if (cases != null && cases.Count > 0)
                {
                    for (int i = 0; i < cases.Count; i++)
                    {
                        if (time[i].Count == 0)
                        {
                            continue;
                        }

                        collectCasesData(i, cases, time, unresponded, responded);
                    }
                }

                if (cases2 != null && cases2.Count > 0)
                {
                    for (int i = 0; i < cases2.Count; i++)
                    {
                        if (time[i].Count == 0)
                        {
                            continue;
                        }

                        collectCasesData(i, cases2, time, unresponded, responded);
                    }
                }

                foreach (string value in unresponded)
                {
                    int i = grid.Rows.Add(value, "尚未填寫");
                    grid.Rows[i].Cells[1].Style.ForeColor = Color.FromArgb(0xDC3C3C);
                }
                
                foreach (string key in responded.Keys)
                {
                    TimeSpan interval = DateTime.Now.Subtract(responded[key]);
                    int i = grid.Rows.Add(key, parseIntervalToWords(interval));

                    if(interval.Days > 3)
                    {
                        grid.Rows[i].Cells[1].Style.ForeColor = Color.FromArgb(0xDC3C3C);
                    } else
                    {
                        grid.Rows[i].Cells[1].Style.ForeColor = Color.FromArgb(0x5548FF);
                    }
                }

            }
            else
            {
                MessageBox.Show("錯誤", "找不到資料!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static string parseIntervalToWords(TimeSpan time)
        {
            string t = "";
            int days = (int) Math.Floor(time.TotalDays);
            int hours = (int)Math.Floor(time.TotalHours);
            int minutes = (int)Math.Floor(time.TotalMinutes);
            if (days > 1)
            {
                t += days + "天";

                if (hours > 1)
                {
                    t += "又" + (days - hours * 24) + "小時前";
                }
            } else
            {
                if (hours > 1)
                {
                    t += hours + "小時";

                    if (minutes > 1)
                    {
                        t += "又" + (minutes - hours * 60) + "分鐘前";
                    }
                }
                else
                {
                    if (minutes > 0)
                    {
                        t += minutes + "分鐘前";
                    } else
                    {
                        t += "剛剛";
                    }
                }
            }
            return t;
        }

        private static void collectCasesData(int i, IList<IList<Object>> cases, IList<IList<Object>> time, List<string> unresponded, Dictionary<string, DateTime> responded)
        {
            if (cases.Count > i)
            {
                if (cases[i].Count > 0)
                {
                    string data = cases[i][0].ToString();
                    foreach (string value_in_list in unresponded.ToArray())
                    {
                        if (value_in_list.Equals(data))
                        {
                            unresponded.Remove(value_in_list);
                            responded.Add(value_in_list, parseTimeByString(time[i][0].ToString()));
                            break;
                        }
                    }

                    if (responded.ContainsKey(data))
                    {
                        if (parseTimeByString(time[i][0].ToString()).CompareTo(responded[data]) > 0)
                        {
                            responded[data] = parseTimeByString(time[i][0].ToString());
                        }
                    }
                }
            }
        }

        private static DateTime parseTimeByString(string time)
        {
            //2022/1/27 上午 7:23:44
            //2022/1/27 上午 7:23:44
            if (time[time.IndexOf("午") + 3] == ':')
            {
                time = time.Insert(time.IndexOf("午") + 2, "0");
            }
            if (time[time.LastIndexOf("/") + 2] == ' ')
            {
                time = time.Insert(time.LastIndexOf("/") + 1, "0");
            }
            if (time[time.IndexOf("/") + 1] != '0')
            {
                time = time.Insert(time.IndexOf("/") + 1, "0");
            }

            Console.WriteLine(time);

            return DateTime.ParseExact(time, "yyyy/MM/dd tt hh:mm:ss", new CultureInfo("zh-TW", false).DateTimeFormat);
        }

        public void downloadFile(string id)
        {


            Excel.Application app = new Excel.Application();

            if (app == null)
            {
                if(MessageBox.Show("未安裝Excel無法執行，是否要直接下載未處理的檔案?", "錯誤", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    string url = "https://docs.google.com/spreadsheets/d/" + id + "/export?format=xlsx";
                    try
                    {
                        System.Diagnostics.Process.Start(url);
                    }
                    catch (System.ComponentModel.Win32Exception noBrowser)
                    {
                        if (noBrowser.ErrorCode == -2147467259)
                            MessageBox.Show(noBrowser.Message);
                    }
                    catch (System.Exception other)
                    {
                        MessageBox.Show(other.Message);
                    }
                }
                return;
            }

            Excel.Workbook workBook;
            workBook = app.Workbooks.Add(true);

            //get data

            UserCredential credential;
            SheetsService service = null;
            try
            {
                using (var stream =
                    new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
                {
                    string credPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal);
                    credPath = Path.Combine(credPath, ".credentials/sheets.googleapis.com-dotnet-quickstart.json");
                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.FromStream(stream).Secrets,
                        Scopes,
                        "user",
                        CancellationToken.None,
                        new FileDataStore(credPath, true)).Result;
                    Console.WriteLine("Credential file saved to: " + credPath);
                }

                if (credential == null)
                {
                    this.Close();
                }

                // Create Google Sheets API service.
                service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show("資料傳輸出現問題!", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            String spreadsheetId = id;
            String range = "表單回應 1!A2:P50000";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId, range);
            IList<IList<object>> datas = request.Execute().Values;
            Dictionary<string, Excel.Worksheet> sheets = new Dictionary<string, Excel.Worksheet>();
            Dictionary<string, int> sheets_last_number = new Dictionary<string, int>();

            // make sheet

            foreach (IList<object> data in datas) {

                if(data[1].Equals(""))
                {
                    continue;
                }

                string case_name;

                if (data[1].ToString().Equals("長輩本人"))
                {
                    case_name = data[2].ToString();
                } else
                {
                    case_name = data[7].ToString();
                }

                Excel.Worksheet workSheet;

                if (sheets.ContainsKey(case_name))
                {
                    workSheet = sheets[case_name];
                } else
                {
                    workSheet = workBook.Worksheets.Add();
                    workSheet.Name = case_name;
                    sheets.Add(case_name, workSheet);
                    workSheet.Cells[1][1] = "填寫者";
                    workSheet.Cells[2][1] = "填寫時間";
                    workSheet.Cells[3][1] = "長輩體溫";
                    workSheet.Cells[4][1] = "主要照顧者體溫";
                    workSheet.Cells[5][1] = "長輩出現疑似COVID-19之症狀";
                    workSheet.Cells[6][1] = "近期接觸或出入之場所";
                    workSheet.Cells[7][1] = "近1個月內的群聚史";
                    sheets_last_number[case_name] = 2;
                }

                int t = sheets_last_number[case_name];
                workSheet.Cells[2][t] = data[0];
                if (data[1].ToString().Equals("長輩本人"))
                {
                    workSheet.Cells[1][t] = "長輩本人";
                    workSheet.Cells[3][t] = data[3];
                    workSheet.Cells[4][t] = "X";
                    workSheet.Cells[5][t] = data[6];
                    workSheet.Cells[6][t] = data[4];
                    workSheet.Cells[7][t] = data[5];
                }
                else
                {
                    workSheet.Cells[1][t] = "家屬/陪同照顧者";
                    workSheet.Cells[3][t] = data[8];
                    workSheet.Cells[4][t] = data[9];
                    workSheet.Cells[5][t] = data[14];
                    workSheet.Cells[6][t] = data[12];
                    workSheet.Cells[7][t] = data[13];
                }

                sheets_last_number[case_name] += 1;

                workSheet.Columns[1].ColumnWidth = 18;
                workSheet.Columns[2].ColumnWidth = 20;
                workSheet.Columns[3].ColumnWidth = 12;
                workSheet.Columns[4].ColumnWidth = 18;
                workSheet.Columns[5].ColumnWidth = 30;
                workSheet.Columns[6].ColumnWidth = 25;
                workSheet.Columns[7].ColumnWidth = 20;
            }

            app.Visible = true;
            workBook.Worksheets["工作表1"].Delete();
        }

        private void generate_Click(object sender, EventArgs e)
        {
            var ResxFile = "SheetResource.resx";
            string sheetid = "";
            string cases = "";
            using (var reader = new ResXResourceReader(ResxFile))
            {
                foreach (DictionaryEntry d in reader)
                {
                    if (d.Key.ToString().Equals("SHEET_ID"))
                    {
                        sheetid = d.Value.ToString();
                    }
                    if (d.Key.ToString().Equals("CASES"))
                    {
                        cases = d.Value.ToString();
                    }
                }
            }
            if (sheetid.Equals(""))
            {
                MessageBox.Show("尚未設定試算表ID!", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else if (cases.Equals(""))
            {
                MessageBox.Show("尚未設定個案資料!", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            Load_Data(sheetid, cases, grid, generate);
            update.Enabled = true;
        }

        private void update_Click(object sender, EventArgs e)
        {
            new input().Show();
        }

        private void download_Click(object sender, EventArgs e)
        {
            var ResxFile = "SheetResource.resx";
            string sheetid = "";
            using (var reader = new ResXResourceReader(ResxFile))
            {
                foreach (DictionaryEntry d in reader)
                {
                    if (d.Key.ToString().Equals("SHEET_ID"))
                    {
                        sheetid = d.Value.ToString();
                    }
                }
            }
            if (sheetid.Equals(""))
            {
                MessageBox.Show("尚未設定試算表ID!", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            downloadFile(sheetid);
        }

        private void grid_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
