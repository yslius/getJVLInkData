using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace getJVLInkData
{
    public partial class Form1 : Form
    {
        private int nDownloadCount;
        private bool JVOpenFlg;
        private clsCodeConv objCodeConv;
        private Timer timer;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string sid = "Test";
            int num = this.AxJVLink1.JVInit(sid);
            if (num != 0)
            {
                MessageBox.Show("JVInit エラー コード：" + num + "：", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                this.Cursor = Cursors.Default;
            }
            this.objCodeConv = new clsCodeConv();
            this.objCodeConv.FileName = System.Windows.Forms.Application.StartupPath + "\\CodeTable.csv";

        }

        private void mnuConfJV_Click(object sender, EventArgs e)
        {
            try
            {
                // リターンコード
                int nReturnCode;

                // 設定画面表示
                nReturnCode = AxJVLink1.JVSetUIProperties();

                if (nReturnCode != 0)
                {
                    MessageBox.Show("JVSetUIPropertiesエラー コード：" + nReturnCode + "：", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.dateTimePicker1.Enabled = true;
            this.rtbData.Text = "";
            CommonOpenFileDialog commonOpenFileDialog = new CommonOpenFileDialog("保存するフォルダを選択してください");
            commonOpenFileDialog.InitialDirectory = Directory.GetCurrentDirectory();
            commonOpenFileDialog.IsFolderPicker = true;
            if (commonOpenFileDialog.ShowDialog() != CommonFileDialogResult.Ok)
                return;
            this.textBox1.Text = commonOpenFileDialog.FileName + "\\";
            string path = commonOpenFileDialog.FileName + "\\01出馬表.csv";
            if (File.Exists(path))
            {
                StreamReader streamReader = new StreamReader(path, Encoding.GetEncoding("shift_jis"));
                bool flag = false;
                while (!flag)
                {
                    string[] strArray = streamReader.ReadLine().Split(',');
                    try
                    {
                        DateTime dateTime = DateTime.Parse(strArray[0]);
                        this.dateTimePicker1.Value = dateTime;
                        this.dateTimePicker1.Enabled = false;
                        this.rtbData.Text = "フォルダ内の出馬表から日付を読み取りました。 " + dateTime.ToLongDateString();
                        flag = true;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine((object)ex);
                    }
                }
                if (flag)
                    return;
                this.rtbData.Text = "フォルダ内の出馬表から日付を読み取れませんでした。";
            }
            else
                this.rtbData.Text = "フォルダ内に出馬表が見つかりませんでした。";

        }

        private void btnGetJVData_Click(object sender, EventArgs e)
        {
            this.button1.Enabled = false;
            this.dateTimePicker1.Enabled = false;
            this.btnGetJVData.Enabled = false;

            int index1 = 0;
            int index2 = 0;
            string str1 = System.Windows.Forms.Application.StartupPath + "\\";
            string str2 = "Template.xlsx";
            string str3 = "Template";
            DateTime datetimeTarg = this.dateTimePicker1.Value;
            string str4 = datetimeTarg.ToString("yyyyMMdd");
            List<string> stringList1 = new List<string>();
            List<List<List<string>>> stringListListList = new List<List<List<string>>>();
            List<string> stringList2 = new List<string>();
            List<List<string>> stringListList1 = new List<List<string>>();
            List<List<cRaceUma>> cRaceUmaListList = new List<List<cRaceUma>>();
            List<cRaceUma> cRaceUmaList = new List<cRaceUma>();
            int num1 = 0;
            int rowWrite = 1;
            OperateCSV operateCsv = new OperateCSV();

            Microsoft.Office.Interop.Excel.Application appExl = null;
            Workbook wbCSV = null;
            Workbook wbTemplate = null;
            Worksheet wsCSV = null;
            Worksheet wsTemplate = null;


            if (this.textBox1.Text == "")
            {
                System.Media.SystemSounds.Asterisk.Play();
                int num2 = (int)MessageBox.Show("保存するフォルダを選択してください。");
                return;
            }

            string text = this.textBox1.Text;
            this.prgDownload.Maximum = 100;
            this.prgDownload.Value = 0;
            if (!this.isRunRace(datetimeTarg))
            {
                System.Media.SystemSounds.Asterisk.Play();
                int num3 = (int)MessageBox.Show("レースが存在しません。");
                return;
            }

            List<string> placeInfoX = this.GetPlaceInfoX(datetimeTarg);
            if (placeInfoX.Count == 0)
            {
                System.Media.SystemSounds.Asterisk.Play();
                int num4 = (int)MessageBox.Show("レースが存在しません。");
                return;
            }

            foreach (string collplace in placeInfoX)
            {
                List<List<string>> raceNumInfoX = this.GetRaceNumInfoX(datetimeTarg, collplace);
                stringListListList.Add(raceNumInfoX);
                foreach (List<string> stringList3 in raceNumInfoX)
                {
                    this.rtbData.Text = "出走馬取得中... " + collplace.Replace("競馬場", "") + Strings.StrConv(stringList3[0], VbStrConv.Wide, 0);
                    List<cRaceUma> raceUmaX = this.GetRaceUmaX(datetimeTarg, collplace, stringList3[0]);
                    cRaceUmaListList.Add(raceUmaX);
                    ++num1;
                    this.prgDownload.Value = 33 * num1 / (12 * placeInfoX.Count);
                }
            }

            this.prgDownload.Value = 33;
            this.rtbData.Text = "調教データ取得中";
            string[] tyoukyouDataAllX = this.GetTyoukyouDataAllX(datetimeTarg);

            appExl = new Microsoft.Office.Interop.Excel.Application();
            appExl.Visible = true;
            wbCSV = appExl.Workbooks.Add(System.Type.Missing);
            wsCSV = wbCSV.ActiveSheet;
            wbTemplate = appExl.Workbooks.Open(Path.GetFullPath(str1 + str2));
            wsTemplate = wbTemplate.Sheets[str3];

            this.prgDownload.Value = 66;
            foreach (List<List<string>> stringListList2 in stringListListList)
            {
                foreach (List<string> stringList3 in stringListList2)
                {
                    this.rtbData.Text = placeInfoX[index2].Replace("競馬場", "") +
                        Strings.StrConv(stringList3[0], VbStrConv.Wide, 0) +
                        ".csv\n" + (object)(index1 + 1) + " / " + (object)num1;

                    // テンプレートシートの値の削除
                    wsTemplate.Cells[1, 1].ClearContents();
                    wsTemplate.Cells[1, 2].ClearContents();
                    Range usedRangeTemp = wsTemplate.UsedRange;
                    object cell1 = wsTemplate.Cells[3, 1];
                    object cell2 = wsTemplate.Cells[usedRangeTemp.Rows.Count, 12];
                    Range rangeTemp = wsTemplate.Range[cell1, cell2];
                    rangeTemp.ClearContents();

                    // 調教データを反映
                    string[,] arrDataTyokyou;
                    double[,] arrdblDataTyokyou;
                    long cntRow = this.PutTyoukyouDataAllX(datetimeTarg, tyoukyouDataAllX,
                        cRaceUmaListList[index1], out arrDataTyokyou, out arrdblDataTyokyou);

                    //for (int i = 0; i < arrdblDataTyokyou.GetLength(0); i++)
                    //{
                    //    // Console.WriteLine(arrdblDataTyokyou[i, 0]);
                    //    if(arrDataTyokyou[i, 1] == "")
                    //    {
                    //        break;
                    //    }
                    //    for (int j = 0; j < 3; j++)
                    //    {
                    //        wsTemplate.Cells[i + 3, j + 1].Value = arrDataTyokyou[i, j];
                    //    }
                    //    for (int j = 3; j < 12; j++)
                    //    {
                    //        wsTemplate.Cells[i + 3, j + 1].Value = arrdblDataTyokyou[i, j - 3];
                    //    }
                    //}

                    cell1 = wsTemplate.Cells[3, 1];
                    cell2 = wsTemplate.Cells[3 + cntRow -1, 3];
                    rangeTemp = wsTemplate.Range[cell1, cell2];
                    rangeTemp.Value = arrDataTyokyou;
                    cell1 = wsTemplate.Cells[3, 4];
                    cell2 = wsTemplate.Cells[3 + cntRow - 1, 12];
                    rangeTemp = wsTemplate.Range[cell1, cell2];
                    rangeTemp.Value = arrdblDataTyokyou;

                    // ファイル名の入力
                    wsTemplate.Cells[1, 1] = "TrainData_" +
                        datetimeTarg.ToString("yyyyMMdd") + "_" +
                        placeInfoX[index2].Replace("競馬場", "") + "_" +
                        stringList3[0];
                    wsTemplate.Cells[1, 2] = stringList3[1];
                    string str5 = placeInfoX[index2].Replace("競馬場", "") +
                        Strings.StrConv(stringList3[0], VbStrConv.Wide, 0) + ".csv";

                    // 小数点の表示
                    usedRangeTemp = wsTemplate.UsedRange;
                    cell1 = wsTemplate.Cells[3, 5];
                    cell2 = wsTemplate.Cells[3 + cntRow - 1, 12];
                    rangeTemp = wsTemplate.Range[cell1, cell2];
                    rangeTemp.NumberFormatLocal = "0.0";

                    // 昇順ソート
                    cell1 = wsTemplate.Cells[3, 1];
                    cell2 = wsTemplate.Cells[3 + cntRow - 1, 12];
                    rangeTemp = wsTemplate.Range[cell1, cell2];
                    rangeTemp.Sort(wsTemplate.Cells[3, 12], XlSortOrder.xlAscending);

                    // 結果の反映
                    cell1 = wsTemplate.Cells[2, 152];
                    cell2 = wsTemplate.Cells[22, 158];
                    rangeTemp = wsTemplate.Range[cell1, cell2];
                    rangeTemp.Copy();
                    wsCSV.Cells[rowWrite, 1].PasteSpecial(XlPasteType.xlPasteValues);
                    rowWrite += 22;

                    appExl.DisplayAlerts = false;
                    wbTemplate.Save();
                    appExl.DisplayAlerts = true;
                    ++index1;
                    this.prgDownload.Value = 66 + 34 * index1 / num1;

                    break;
                }
                ++index2;
                break;
            }

            string str6 = "c" + str4 + ".csv";
            appExl.DisplayAlerts = false;
            wbCSV.SaveAs(text + str6, 6);
            wbCSV.Close(System.Type.Missing, System.Type.Missing, System.Type.Missing);
            wbTemplate.Close(System.Type.Missing, System.Type.Missing, System.Type.Missing);
            appExl.DisplayAlerts = true;

            appExl.Quit();
            Marshal.ReleaseComObject(wsTemplate);
            Marshal.ReleaseComObject(wbTemplate);
            Marshal.ReleaseComObject(wsCSV);
            Marshal.ReleaseComObject(wbCSV);
            Marshal.ReleaseComObject(appExl);

            this.rtbData.Text = str4 + " 調教データ取得完了しました。";
            this.AxJVLink1.JVClose();
            System.Media.SystemSounds.Asterisk.Play();
            this.prgDownload.Value = 100;

            this.button1.Enabled = true;
            this.dateTimePicker1.Enabled = true;
            this.btnGetJVData.Enabled = true;
        }

        private List<string> GetPlaceInfoX(DateTime datetimeTarg)
        {
            List<string> stringList = new List<string>();
            try
            {
                int size = 110000;
                int count = 256;
                JVData_Struct.JV_YS_SCHEDULE jvYsSchedule = new JVData_Struct.JV_YS_SCHEDULE();
                this.tmrDownload.Enabled = false;
                this.prgJVRead.Value = 0;
                string dataspec = "YSCH";
                TimeSpan timeSpan = new TimeSpan(1, 0, 0, 0);
                string str = (datetimeTarg - timeSpan).ToString("yyyyMMdd");
                int year1 = DateTime.Now.Year;
                int year2 = datetimeTarg.AddYears(1).Year;
                int readcount = 0;
                int downloadcount = 0;
                int num1 = this.AxJVLink1.JVOpen(dataspec, str + "000000", 1, ref readcount, ref downloadcount, out string _);
                if (num1 != 0)
                {
                    int num2 = (int)MessageBox.Show("JVOpen エラー：" + (object)num1);
                }
                else
                {
                    this.prgJVRead.Maximum = readcount;
                    if (readcount > 0)
                    {
                        bool flag = false;
                        do
                        {
                            System.Windows.Forms.Application.DoEvents();
                            string buff = new string(char.MinValue, size);
                            string filename = new string(char.MinValue, count);
                            switch (this.AxJVLink1.JVRead(out buff, out size, out filename))
                            {
                                case -503:
                                    int num3 = (int)MessageBox.Show(filename + "が存在しません。");
                                    flag = true;
                                    goto case -3;
                                case -203:
                                    int num4 = (int)MessageBox.Show("JVOpen が行われていません。");
                                    flag = true;
                                    goto case -3;
                                case -201:
                                    int num5 = (int)MessageBox.Show("JVInit が行われていません。");
                                    flag = true;
                                    goto case -3;
                                case -3:
                                    continue;
                                case -1:
                                    ++this.prgJVRead.Value;
                                    goto case -3;
                                case 0:
                                    this.prgJVRead.Value = this.prgJVRead.Maximum;
                                    flag = true;
                                    goto case -3;
                                default:
                                    if (buff.Substring(0, 2) == "YS")
                                    {
                                        jvYsSchedule.SetDataB(ref buff);
                                        DateTime dateTime = DateTime.Parse((jvYsSchedule.id.Year + jvYsSchedule.id.MonthDay).Insert(4, "/").Insert(7, "/"));
                                        int num6 = dateTime > datetimeTarg ? 1 : 0;
                                        string codeName = this.objCodeConv.GetCodeName("2001", jvYsSchedule.id.JyoCD, (short)1);
                                        if (dateTime.Date == datetimeTarg.Date)
                                        {
                                            stringList.Add(codeName);
                                            goto case -3;
                                        }
                                        else
                                            goto case -3;
                                    }
                                    else
                                        goto case -3;
                            }
                        }
                        while (!flag);
                    }
                }
            }
            catch (Exception ex)
            {
                return (List<string>)null;
            }
            int num7 = this.AxJVLink1.JVClose();
            if (num7 != 0)
            {
                int num8 = (int)MessageBox.Show("JVClose エラー：" + (object)num7);
            }
            this.prgJVRead.Value = this.prgJVRead.Maximum;
            return stringList;
        }

        private List<List<string>> GetRaceNumInfoX(DateTime datetimeTarg, string collplace)
        {
            List<List<string>> stringListList = new List<List<string>>();
            List<string> stringList = new List<string>();
            try
            {
                int size = 110000;
                int count = 256;
                JVData_Struct.JV_RA_RACE jvRaRace = new JVData_Struct.JV_RA_RACE();
                this.tmrDownload.Enabled = false;
                this.prgJVRead.Value = 0;
                string dataspec = "RACE";
                TimeSpan timeSpan = new TimeSpan(1, 0, 0, 0);
                string str = (datetimeTarg - timeSpan).ToString("yyyyMMdd");
                int option = DateTime.Now.Year > datetimeTarg.AddYears(1).Year ? 4 : 1;
                int readcount = 0;
                int downloadcount = 0;
                int num1 = this.AxJVLink1.JVOpen(dataspec, str + "000000", option, ref readcount, ref downloadcount, out string _);
                if (num1 != 0)
                {
                    int num2 = (int)MessageBox.Show("JVOpen エラー：" + (object)num1);
                }
                else
                {
                    this.prgJVRead.Maximum = readcount;
                    if (readcount > 0)
                    {
                        bool flag1 = false;
                        bool flag2 = false;
                        do
                        {
                            System.Windows.Forms.Application.DoEvents();
                            string buff = new string(char.MinValue, size);
                            string filename = new string(char.MinValue, count);
                            switch (this.AxJVLink1.JVRead(out buff, out size, out filename))
                            {
                                case -503:
                                    int num3 = (int)MessageBox.Show(filename + "が存在しません。");
                                    flag1 = true;
                                    goto case -3;
                                case -203:
                                    int num4 = (int)MessageBox.Show("JVOpen が行われていません。");
                                    flag1 = true;
                                    goto case -3;
                                case -201:
                                    int num5 = (int)MessageBox.Show("JVInit が行われていません。");
                                    flag1 = true;
                                    goto case -3;
                                case -3:
                                    continue;
                                case -1:
                                    ++this.prgJVRead.Value;
                                    goto case -3;
                                case 0:
                                    this.prgJVRead.Value = this.prgJVRead.Maximum;
                                    flag1 = true;
                                    goto case -3;
                                default:
                                    if (buff.Substring(0, 2) == "RA")
                                    {
                                        jvRaRace.SetDataB(ref buff);
                                        string s = (jvRaRace.id.Year + jvRaRace.id.MonthDay).Insert(4, "/").Insert(7, "/");
                                        DateTime dateTime = DateTime.Parse(s);
                                        if (flag2 && dateTime > datetimeTarg)
                                            flag1 = true;
                                        if (jvRaRace.head.DataKubun != "9" && s == datetimeTarg.ToString("yyyy/MM/dd"))
                                        {
                                            string codeName = this.objCodeConv.GetCodeName("2001", jvRaRace.id.JyoCD, (short)1);
                                            if (collplace == codeName)
                                            {
                                                flag2 = true;
                                                stringList.Add(jvRaRace.id.RaceNum);
                                                stringList.Add(jvRaRace.Kyori);
                                                stringListList.Add(stringList);
                                                stringList = new List<string>();
                                                goto case -3;
                                            }
                                            else
                                                goto case -3;
                                        }
                                        else
                                            goto case -3;
                                    }
                                    else
                                    {
                                        this.AxJVLink1.JVSkip();
                                        goto case -3;
                                    }
                            }
                        }
                        while (!flag1);
                    }
                }
            }
            catch (Exception ex)
            {
                return (List<List<string>>)null;
            }
            int num6 = this.AxJVLink1.JVClose();
            if (num6 != 0)
            {
                int num7 = (int)MessageBox.Show("JVClose エラー：" + (object)num6);
            }
            this.prgJVRead.Value = this.prgJVRead.Maximum;
            return stringListList;
        }

        private List<cRaceUma> GetRaceUmaX(DateTime datetimeTarg, string collplace, string collRace)
        {
            List<cRaceUma> cRaceUmaList = new List<cRaceUma>();
            try
            {
                int size = 110000;
                int count = 256;
                JVData_Struct.JV_SE_RACE_UMA jvSeRaceUma = new JVData_Struct.JV_SE_RACE_UMA();
                this.tmrDownload.Enabled = false;
                this.prgJVRead.Value = 0;
                string dataspec = "RACE";
                TimeSpan timeSpan = new TimeSpan(1, 0, 0, 0);
                string str = (datetimeTarg - timeSpan).ToString("yyyyMMdd");
                int option = DateTime.Now.Year > datetimeTarg.AddYears(1).Year ? 4 : 1;
                int readcount = 0;
                int downloadcount = 0;
                int num1 = this.AxJVLink1.JVOpen(dataspec, str + "000000", option, ref readcount, ref downloadcount, out string _);
                if (num1 != 0)
                {
                    int num2 = (int)MessageBox.Show("JVOpen エラー：" + (object)num1);
                }
                else
                {
                    this.prgJVRead.Maximum = readcount;
                    if (readcount > 0)
                    {
                        bool flag1 = false;
                        bool flag2 = false;
                        do
                        {
                            System.Windows.Forms.Application.DoEvents();
                            string buff = new string(char.MinValue, size);
                            string filename = new string(char.MinValue, count);
                            switch (this.AxJVLink1.JVRead(out buff, out size, out filename))
                            {
                                case -503:
                                    int num3 = (int)MessageBox.Show(filename + "が存在しません。");
                                    flag1 = true;
                                    goto case -3;
                                case -203:
                                    int num4 = (int)MessageBox.Show("JVOpen が行われていません。");
                                    flag1 = true;
                                    goto case -3;
                                case -201:
                                    int num5 = (int)MessageBox.Show("JVInit が行われていません。");
                                    flag1 = true;
                                    goto case -3;
                                case -3:
                                    continue;
                                case -1:
                                    ++this.prgJVRead.Value;
                                    goto case -3;
                                case 0:
                                    this.prgJVRead.Value = this.prgJVRead.Maximum;
                                    flag1 = true;
                                    goto case -3;
                                default:
                                    if (buff.Substring(0, 2) == "SE")
                                    {
                                        jvSeRaceUma.SetDataB(ref buff);
                                        string s = (jvSeRaceUma.id.Year + jvSeRaceUma.id.MonthDay).Insert(4, "/").Insert(7, "/");
                                        DateTime dateTime = DateTime.Parse(s);
                                        if (flag2 && dateTime > datetimeTarg)
                                            flag1 = true;
                                        if (s == datetimeTarg.ToString("yyyy/MM/dd"))
                                        {
                                            string codeName = this.objCodeConv.GetCodeName("2001", jvSeRaceUma.id.JyoCD, (short)1);
                                            if (collplace == codeName && collRace == jvSeRaceUma.id.RaceNum)
                                            {
                                                flag2 = true;
                                                cRaceUmaList.Add(new cRaceUma()
                                                {
                                                    strdate = jvSeRaceUma.id.Year + jvSeRaceUma.id.MonthDay,
                                                    nameJyo = codeName,
                                                    racenum = jvSeRaceUma.id.RaceNum,
                                                    KettoNum = jvSeRaceUma.KettoNum,
                                                    Bamei = jvSeRaceUma.Bamei.Trim(),
                                                    Umaban = jvSeRaceUma.Umaban
                                                });
                                                goto case -3;
                                            }
                                            else
                                                goto case -3;
                                        }
                                        else
                                            goto case -3;
                                    }
                                    else
                                    {
                                        this.AxJVLink1.JVSkip();
                                        goto case -3;
                                    }
                            }
                        }
                        while (!flag1);
                    }
                }
            }
            catch (Exception ex)
            {
                return (List<cRaceUma>)null;
            }
            int num6 = this.AxJVLink1.JVClose();
            if (num6 != 0)
            {
                int num7 = (int)MessageBox.Show("JVClose エラー：" + (object)num6);
            }
            this.prgJVRead.Value = this.prgJVRead.Maximum;
            return cRaceUmaList;
        }

        private string[] GetTyoukyouDataAllX(DateTime datetimeTarg)
        {
            string[] strArray = new string[300000];
            int index = 0;
            try
            {
                int size = 110000;
                int count = 256;
                this.tmrDownload.Enabled = false;
                this.prgJVRead.Value = 0;
                string dataspec = "SLOP";
                TimeSpan timeSpan1 = new TimeSpan(180, 0, 0, 0);
                DateTime dateTime1 = datetimeTarg - timeSpan1;
                string str = dateTime1.ToString("yyyyMMdd");
                dateTime1 = DateTime.Now;
                int year1 = dateTime1.Year;
                dateTime1 = datetimeTarg.AddYears(1);
                int year2 = dateTime1.Year;
                int option = year1 > year2 ? 4 : 1;
                int readcount = 0;
                int downloadcount = 0;
                int num1 = this.AxJVLink1.JVOpen(dataspec, str + "000000", option, ref readcount, ref downloadcount, out string _);
                if (num1 != 0)
                {
                    int num2 = (int)MessageBox.Show("JVOpen エラー：" + (object)num1);
                    Environment.Exit(0);
                }
                this.prgJVRead.Maximum = readcount;
                if (readcount == 0)
                {
                    int num2 = (int)MessageBox.Show("nReadCount == 0");
                    Environment.Exit(0);
                }
                bool flag1 = false;
                bool flag2 = true;
                do
                {
                    System.Windows.Forms.Application.DoEvents();
                    string buff = new string(char.MinValue, size);
                    string filename = new string(char.MinValue, count);
                    switch (this.AxJVLink1.JVRead(out buff, out size, out filename))
                    {
                        case -503:
                            int num2 = (int)MessageBox.Show(filename + "が存在しません。");
                            flag1 = true;
                            goto case -3;
                        case -203:
                            int num3 = (int)MessageBox.Show("JVOpen が行われていません。");
                            flag1 = true;
                            goto case -3;
                        case -201:
                            int num4 = (int)MessageBox.Show("JVInit が行われていません。");
                            flag1 = true;
                            goto case -3;
                        case -3:
                            continue;
                        case -1:
                            if (this.prgJVRead.Value + 1 > this.prgJVRead.Maximum)
                                this.prgJVRead.Maximum = this.prgJVRead.Value + 1;
                            ++this.prgJVRead.Value;
                            this.rtbData.Text = string.Format("調教データ取得中 {0}/{1}", (object)this.prgJVRead.Value, (object)this.prgJVRead.Maximum);
                            goto case -3;
                        case 0:
                            this.prgJVRead.Value = this.prgJVRead.Maximum;
                            flag1 = true;
                            goto case -3;
                        default:
                            DateTime dateTime2 = DateTime.Parse(buff.Substring(12, 8).Insert(4, "/").Insert(7, "/"));
                            if (flag2)
                            {
                                flag2 = false;
                                TimeSpan timeSpan2 = datetimeTarg - dateTime2;
                                double totalDays1 = timeSpan2.TotalDays;
                                timeSpan2 = DateTime.Today - datetimeTarg;
                                double totalDays2 = timeSpan2.TotalDays;
                                this.prgJVRead.Maximum = (int)((double)readcount * (totalDays1 / (totalDays1 + totalDays2)));
                            }
                            if (dateTime2 > datetimeTarg)
                                flag1 = true;
                            strArray[index] = buff;
                            ++index;
                            goto case -3;
                    }
                }
                while (!flag1);
            }
            catch (Exception ex)
            {
                return (string[])null;
            }
            int num5 = this.AxJVLink1.JVClose();
            if (num5 != 0)
            {
                int num6 = (int)MessageBox.Show("JVClose エラー：" + (object)num5);
            }
            this.prgJVRead.Value = this.prgJVRead.Maximum;
            return strArray;
        }

        private long PutTyoukyouDataAllX(DateTime datetimeTarg, string[] dataTyokyo, 
            List<cRaceUma> ListUmas, out string[,] arrDataTyokyou, out double[,] arrdblDataTyokyou)
        {
            arrDataTyokyou = new string[2000, 3];
            arrdblDataTyokyou = new double[2000, 9];
            int index1 = 0;
            for (int index2 = 0; index2 < dataTyokyo.Length && dataTyokyo[index2] != null; ++index2)
            {
                if (dataTyokyo[index2].Substring(55, 3) != "000")
                {
                    foreach (cRaceUma listUma in ListUmas)
                    {
                        if (listUma.KettoNum == dataTyokyo[index2].Substring(24, 10))
                        {
                            JVData_Struct.JV_HC_HANRO jvHcHanro = new JVData_Struct.JV_HC_HANRO();
                            jvHcHanro.SetDataB(ref dataTyokyo[index2]);
                            string TresenKubun = (int.Parse(jvHcHanro.TresenKubun) + 1).ToString();
                            string codeName = this.objCodeConv.GetCodeName("2301",TresenKubun, (short)2);
                            arrDataTyokyou[index1, 0] = codeName;
                            arrDataTyokyou[index1, 1] = jvHcHanro.ChokyoDate.Year + jvHcHanro.ChokyoDate.Month + jvHcHanro.ChokyoDate.Day;
                            string bamei = listUma.Bamei;
                            int num = int.Parse(listUma.Umaban);
                            arrDataTyokyou[index1, 2] = bamei;
                            arrdblDataTyokyou[index1, 0] = (double)num;
                            arrdblDataTyokyou[index1, 1] = double.Parse(jvHcHanro.HaronTime4) / 10.0;
                            arrdblDataTyokyou[index1, 2] = double.Parse(jvHcHanro.HaronTime3) / 10.0;
                            arrdblDataTyokyou[index1, 3] = double.Parse(jvHcHanro.HaronTime2) / 10.0;
                            arrdblDataTyokyou[index1, 4] = double.Parse(jvHcHanro.LapTime1) / 10.0;
                            arrdblDataTyokyou[index1, 5] = double.Parse(jvHcHanro.LapTime4) / 10.0;
                            arrdblDataTyokyou[index1, 6] = double.Parse(jvHcHanro.LapTime3) / 10.0;
                            arrdblDataTyokyou[index1, 7] = double.Parse(jvHcHanro.LapTime2) / 10.0;
                            arrdblDataTyokyou[index1, 8] = double.Parse(jvHcHanro.LapTime1) / 10.0;
                            ++index1;
                        }
                    }
                }
            }
            return index1;
        }


        private bool isRunRace(DateTime datetimeTarg)
        {
            bool flag1 = false;
            try
            {
                int size = 110000;
                int count = 256;
                JVData_Struct.JV_YS_SCHEDULE jvYsSchedule = new JVData_Struct.JV_YS_SCHEDULE();
                this.tmrDownload.Enabled = false;
                this.prgJVRead.Value = 0;
                string dataspec = "YSCH";
                TimeSpan timeSpan = new TimeSpan(4, 0, 0, 0);
                string str = (datetimeTarg - timeSpan).ToString("yyyyMMdd");
                int year1 = DateTime.Now.Year;
                int year2 = datetimeTarg.AddYears(1).Year;
                int readcount = 0;
                int downloadcount = 0;
                int num1 = this.AxJVLink1.JVOpen(dataspec, str + "000000", 1, ref readcount, ref downloadcount, out string _);
                if (num1 != 0)
                {
                    int num2 = (int)MessageBox.Show("JVOpen エラー：" + (object)num1);
                }
                else
                {
                    this.prgJVRead.Maximum = readcount;
                    if (readcount > 0)
                    {
                        bool flag2 = false;
                        do
                        {
                            System.Windows.Forms.Application.DoEvents();
                            string buff = new string(char.MinValue, size);
                            string filename = new string(char.MinValue, count);
                            switch (this.AxJVLink1.JVRead(out buff, out size, out filename))
                            {
                                case -503:
                                    int num3 = (int)MessageBox.Show(filename + "が存在しません。");
                                    flag2 = true;
                                    goto case -3;
                                case -203:
                                    int num4 = (int)MessageBox.Show("JVOpen が行われていません。");
                                    flag2 = true;
                                    goto case -3;
                                case -201:
                                    int num5 = (int)MessageBox.Show("JVInit が行われていません。");
                                    flag2 = true;
                                    goto case -3;
                                case -3:
                                    continue;
                                case -1:
                                    ++this.prgJVRead.Value;
                                    goto case -3;
                                case 0:
                                    this.prgJVRead.Value = this.prgJVRead.Maximum;
                                    flag2 = true;
                                    goto case -3;
                                default:
                                    jvYsSchedule.SetDataB(ref buff);
                                    if (jvYsSchedule.id.Year + jvYsSchedule.id.MonthDay == datetimeTarg.ToString("yyyyMMdd"))
                                    {
                                        flag1 = true;
                                        flag2 = true;
                                        goto case -3;
                                    }
                                    else
                                        goto case -3;
                            }
                        }
                        while (!flag2);
                    }
                }
            }
            catch (Exception ex)
            {
            }
            int num6 = this.AxJVLink1.JVClose();
            if (num6 != 0)
            {
                int num7 = (int)MessageBox.Show("JVClose エラー：" + (object)num6);
            }
            if (flag1)
                return true;
            int num8 = (int)MessageBox.Show("選択した日付の開催はありません。");
            return false;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            double[,] arrdblDataTyokyou;
            arrdblDataTyokyou = new double[2000, 9];
            Console.WriteLine(arrdblDataTyokyou.GetLength(1));


            string str1 = System.Windows.Forms.Application.StartupPath + "\\";
            string str2 = "Template.xlsx";
            string str3 = "Template";
            DateTime datetimeTarg = this.dateTimePicker1.Value;
            string str4 = datetimeTarg.ToString("yyyyMMdd");

            Microsoft.Office.Interop.Excel.Application appExl = null;
            Workbook workbook1 = null;
            Workbook workbook2 = null;
            Worksheet worksheet1 = null;
            Worksheet worksheet2 = null;

            appExl = new Microsoft.Office.Interop.Excel.Application();
            appExl.Visible = true;

            workbook1 = appExl.Workbooks.Add(System.Type.Missing);
            worksheet1 = workbook1.ActiveSheet;
            workbook2 = appExl.Workbooks.Open(Path.GetFullPath(str1 + str2),
                System.Type.Missing, System.Type.Missing, System.Type.Missing,
                System.Type.Missing, System.Type.Missing, System.Type.Missing,
                System.Type.Missing, System.Type.Missing, System.Type.Missing,
                System.Type.Missing, System.Type.Missing, System.Type.Missing,
                System.Type.Missing, System.Type.Missing);
            worksheet2 = workbook2.Sheets[str3];

            Range usedRange2 = worksheet2.UsedRange;
            object cell1 = worksheet2.Cells[2, 152];
            object cell2 = worksheet2.Cells[22, 158];
            Range range2 = worksheet2.Range[cell1, cell2];

            int WriteR = 1;
            worksheet2.Copy(worksheet1);
            worksheet1 = workbook1.ActiveSheet;
            range2.Copy(worksheet1.Cells[WriteR, 1]);

            string str6 = "c" + str4 + ".csv";
            string text = this.textBox1.Text;

            appExl.DisplayAlerts = false;
            workbook1.SaveAs(text + str6, 6);
            workbook1.Close(System.Type.Missing, System.Type.Missing, System.Type.Missing);
            workbook2.Close(System.Type.Missing, System.Type.Missing, System.Type.Missing);
            appExl.DisplayAlerts = true;

            appExl.Quit();
            Marshal.ReleaseComObject(worksheet2);
            Marshal.ReleaseComObject(workbook2);
            //Marshal.ReleaseComObject(worksheet1);
            //Marshal.ReleaseComObject(workbook1);
            Marshal.ReleaseComObject(appExl);
        }
    }
}
