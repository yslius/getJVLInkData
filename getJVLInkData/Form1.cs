using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.WindowsAPICodePack.Dialogs;
using System.IO;
using Microsoft.Office.Interop.Excel;

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
            OperateCSV operateCsv = new OperateCSV();

            Microsoft.Office.Interop.Excel.Application application = null;
            Workbook workbook = null;
            Workbook workbook2 = null;
            Worksheet worksheet = null;


            if (this.textBox1.Text == "")
            {
                int num2 = (int)MessageBox.Show("保存するフォルダを選択してください。");
            }
            else
            {
                string text = this.textBox1.Text;
                this.prgDownload.Maximum = 100;
                this.prgDownload.Value = 0;
                if (!this.isRunRace(datetimeTarg))
                {
                    int num3 = (int)MessageBox.Show("レースが存在しません。");
                }
                else
                {
                    List<string> placeInfoX = this.GetPlaceInfoX(datetimeTarg);
                    if (placeInfoX.Count == 0)
                    {
                        int num4 = (int)MessageBox.Show("レースが存在しません。");
                    }
                    else
                    {
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
                        // ISSUE: variable of a compiler-generated type
                        Microsoft.Office.Interop.Excel.Application instance = (Microsoft.Office.Interop.Excel.Application)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("00024500-0000-0000-C000-000000000046")));
                        instance.Visible = false;

                        Workbook workbook1 = instance.Workbooks.Add(System.Type.Missing);
                        Worksheet worksheet1 = Form1.\u003C\u003Eo__11.\u003C\u003Ep__0.Target((CallSite)Form1.\u003C\u003Eo__11.\u003C\u003Ep__0, workbook1.ActiveSheet);
                        // ISSUE: reference to a compiler-generated method
                        // ISSUE: variable of a compiler-generated type
                        Workbook workbook2 = instance.Workbooks.Open(Path.GetFullPath(str1 + str2), System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing);
                        // ISSUE: reference to a compiler-generated field
                        // ISSUE: reference to a compiler-generated field
                        // ISSUE: variable of a compiler-generated type
                        Worksheet worksheet2 = Form1.\u003C\u003Eo__11.\u003C\u003Ep__1.Target((CallSite)Form1.\u003C\u003Eo__11.\u003C\u003Ep__1, workbook2.Sheets[(object)str3]);
                        this.prgDownload.Value = 66;
                        foreach (List<List<string>> stringListList2 in stringListListList)
                        {
                            foreach (List<string> stringList3 in stringListList2)
                            {
                                this.rtbData.Text = placeInfoX[index2].Replace("競馬場", "") + Strings.StrConv(stringList3[0], VbStrConv.Wide, 0) + ".csv\n" + (object)(index1 + 1) + " / " + (object)num1;
                                // ISSUE: variable of a compiler-generated type
                                Range usedRange1 = worksheet2.UsedRange;
                                // ISSUE: reference to a compiler-generated field
                                if (Form1.\u003C\u003Eo__11.\u003C\u003Ep__4 == null)
                                {
                                    // ISSUE: reference to a compiler-generated field
                                    Form1.\u003C\u003Eo__11.\u003C\u003Ep__4 = CallSite<Func<CallSite, object, Range>>.Create(Binder.Convert(CSharpBinderFlags.None, typeof(Range), typeof(Form1)));
                                }
                                // ISSUE: reference to a compiler-generated field
                                Func<CallSite, object, Range> target1 = Form1.\u003C\u003Eo__11.\u003C\u003Ep__4.Target;
                                // ISSUE: reference to a compiler-generated field
                                CallSite<Func<CallSite, object, Range>> p4 = Form1.\u003C\u003Eo__11.\u003C\u003Ep__4;
                                // ISSUE: reference to a compiler-generated field
                                if (Form1.\u003C\u003Eo__11.\u003C\u003Ep__3 == null)
                                {
                                    // ISSUE: reference to a compiler-generated field
                                    Form1.\u003C\u003Eo__11.\u003C\u003Ep__3 = CallSite<Func<CallSite, object, object, object, object>>.Create(Binder.GetIndex(CSharpBinderFlags.None, typeof(Form1), (IEnumerable<CSharpArgumentInfo>)new CSharpArgumentInfo[3]
                                    {
                                        CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
                                        CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
                                        CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
                                    }));
                                }
                                // ISSUE: reference to a compiler-generated field
                                Func<CallSite, object, object, object, object> target2 = Form1.\u003C\u003Eo__11.\u003C\u003Ep__3.Target;
                                // ISSUE: reference to a compiler-generated field
                                CallSite<Func<CallSite, object, object, object, object>> p3 = Form1.\u003C\u003Eo__11.\u003C\u003Ep__3;
                                // ISSUE: reference to a compiler-generated field
                                if (Form1.\u003C\u003Eo__11.\u003C\u003Ep__2 == null)
                                {
                                    // ISSUE: reference to a compiler-generated field
                                    Form1.\u003C\u003Eo__11.\u003C\u003Ep__2 = CallSite<Func<CallSite, Worksheet, object>>.Create(Binder.GetMember(CSharpBinderFlags.ResultIndexed, "Range", typeof(Form1), (IEnumerable<CSharpArgumentInfo>)new CSharpArgumentInfo[1]
                                    {
                                        CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType, (string) null)
                                    }));
                                }
                                // ISSUE: reference to a compiler-generated field
                                // ISSUE: reference to a compiler-generated field
                                object obj1 = Form1.\u003C\u003Eo__11.\u003C\u003Ep__2.Target((CallSite)Form1.\u003C\u003Eo__11.\u003C\u003Ep__2, worksheet2);
                                object cell1 = worksheet2.Cells[(object)3, (object)1];
                                object cell2 = worksheet2.Cells[(object)usedRange1.Rows.Count, (object)12];
                                object obj2 = target2((CallSite)p3, obj1, cell1, cell2);
                                // ISSUE: variable of a compiler-generated type
                                Range range1 = target1((CallSite)p4, obj2);
                                // ISSUE: reference to a compiler-generated method
                                range1.ClearContents();
                                string[,] arrDataTyokyou;
                                double[,] arrdblDataTyokyou;
                                this.PutTyoukyouDataAllX(datetimeTarg, tyoukyouDataAllX, cRaceUmaListList[index1], out arrDataTyokyou, out arrdblDataTyokyou);
                                worksheet2.Cells[(object)1, (object)1] = (object)("TrainData_" + datetimeTarg.ToString("yyyyMMdd") + "_" + placeInfoX[index2].Replace("競馬場", "") + "_" + stringList3[0]);
                                worksheet2.Cells[(object)1, (object)2] = (object)stringList3[1];
                                string str5 = placeInfoX[index2].Replace("競馬場", "") + Strings.StrConv(stringList3[0], VbStrConv.Wide, 0) + ".csv";
                                int num5 = 0;
                                for (int index3 = 0; index3 < arrDataTyokyou.GetLength(0); ++index3)
                                {
                                    if (arrDataTyokyou[index3, 0] == null)
                                    {
                                        num5 = index3;
                                        break;
                                    }
                                }
                                // ISSUE: reference to a compiler-generated field
                                if (Form1.\u003C\u003Eo__11.\u003C\u003Ep__7 == null)
                                {
                                    // ISSUE: reference to a compiler-generated field
                                    Form1.\u003C\u003Eo__11.\u003C\u003Ep__7 = CallSite<Func<CallSite, object, Range>>.Create(Binder.Convert(CSharpBinderFlags.None, typeof(Range), typeof(Form1)));
                                }
                                // ISSUE: reference to a compiler-generated field
                                Func<CallSite, object, Range> target3 = Form1.\u003C\u003Eo__11.\u003C\u003Ep__7.Target;
                                // ISSUE: reference to a compiler-generated field
                                CallSite<Func<CallSite, object, Range>> p7 = Form1.\u003C\u003Eo__11.\u003C\u003Ep__7;
                                // ISSUE: reference to a compiler-generated field
                                if (Form1.\u003C\u003Eo__11.\u003C\u003Ep__6 == null)
                                {
                                    // ISSUE: reference to a compiler-generated field
                                    Form1.\u003C\u003Eo__11.\u003C\u003Ep__6 = CallSite<Func<CallSite, object, object, object, object>>.Create(Binder.GetIndex(CSharpBinderFlags.None, typeof(Form1), (IEnumerable<CSharpArgumentInfo>)new CSharpArgumentInfo[3]
                                    {
                                        CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
                                        CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
                                        CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
                                    }));
                                }
                                // ISSUE: reference to a compiler-generated field
                                Func<CallSite, object, object, object, object> target4 = Form1.\u003C\u003Eo__11.\u003C\u003Ep__6.Target;
                                // ISSUE: reference to a compiler-generated field
                                CallSite<Func<CallSite, object, object, object, object>> p6 = Form1.\u003C\u003Eo__11.\u003C\u003Ep__6;
                                    // ISSUE: reference to a compiler-generated field
                                if (Form1.\u003C\u003Eo__11.\u003C\u003Ep__5 == null)
                                {
                                    // ISSUE: reference to a compiler-generated field
                                    Form1.\u003C\u003Eo__11.\u003C\u003Ep__5 = CallSite<Func<CallSite, Worksheet, object>>.Create(Binder.GetMember(CSharpBinderFlags.ResultIndexed, "Range", typeof(Form1), (IEnumerable<CSharpArgumentInfo>)new CSharpArgumentInfo[1]
                                    {
                                        CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType, (string) null)
                                    }));
                                }
                                // ISSUE: reference to a compiler-generated field
                                // ISSUE: reference to a compiler-generated field
                                object obj3 = Form1.\u003C\u003Eo__11.\u003C\u003Ep__5.Target((CallSite)Form1.\u003C\u003Eo__11.\u003C\u003Ep__5, worksheet2);
                                object cell3 = worksheet2.Cells[(object)3, (object)1];
                                object cell4 = worksheet2.Cells[(object)(num5 + 2), (object)3];
                                object obj4 = target4((CallSite)p6, obj3, cell3, cell4);
                                // ISSUE: variable of a compiler-generated type
                                Range range2 = target3((CallSite)p7, obj4);
                                // ISSUE: reference to a compiler-generated method
                                range2.set_Value(System.Type.Missing, (object)arrDataTyokyou);
                                // ISSUE: reference to a compiler-generated field
                                if (Form1.\u003C\u003Eo__11.\u003C\u003Ep__10 == null)
                {
                  // ISSUE: reference to a compiler-generated field
                  Form1.\u003C\u003Eo__11.\u003C\u003Ep__10 = CallSite<Func<CallSite, object, Range>>.Create(Binder.Convert(CSharpBinderFlags.None, typeof (Range), typeof (Form1)));
                }
                // ISSUE: reference to a compiler-generated field
                Func<CallSite, object, Range> target5 = Form1.\u003C\u003Eo__11.\u003C\u003Ep__10.Target;
                // ISSUE: reference to a compiler-generated field
                CallSite<Func<CallSite, object, Range>> p10 = Form1.\u003C\u003Eo__11.\u003C\u003Ep__10;
                // ISSUE: reference to a compiler-generated field
                if (Form1.\u003C\u003Eo__11.\u003C\u003Ep__9 == null)
                {
                  // ISSUE: reference to a compiler-generated field
                  Form1.\u003C\u003Eo__11.\u003C\u003Ep__9 = CallSite<Func<CallSite, object, object, object, object>>.Create(Binder.GetIndex(CSharpBinderFlags.None, typeof (Form1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[3]
                  {
                    CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
                    CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
                    CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
                  }));
                }
                // ISSUE: reference to a compiler-generated field
                Func<CallSite, object, object, object, object> target6 = Form1.\u003C\u003Eo__11.\u003C\u003Ep__9.Target;
                // ISSUE: reference to a compiler-generated field
                CallSite<Func<CallSite, object, object, object, object>> p9 = Form1.\u003C\u003Eo__11.\u003C\u003Ep__9;
                // ISSUE: reference to a compiler-generated field
                if (Form1.\u003C\u003Eo__11.\u003C\u003Ep__8 == null)
                {
                  // ISSUE: reference to a compiler-generated field
                  Form1.\u003C\u003Eo__11.\u003C\u003Ep__8 = CallSite<Func<CallSite, Worksheet, object>>.Create(Binder.GetMember(CSharpBinderFlags.ResultIndexed, "Range", typeof (Form1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
                  {
                    CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType, (string) null)
                  }));
                }
                // ISSUE: reference to a compiler-generated field
                // ISSUE: reference to a compiler-generated field
                object obj5 = Form1.\u003C\u003Eo__11.\u003C\u003Ep__8.Target((CallSite) Form1.\u003C\u003Eo__11.\u003C\u003Ep__8, worksheet2);
                object cell5 = worksheet2.Cells[(object) 3, (object) 4];
                object cell6 = worksheet2.Cells[(object) (num5 + 2), (object) 12];
                object obj6 = target6((CallSite) p9, obj5, cell5, cell6);
                // ISSUE: variable of a compiler-generated type
                Range range3 = target5((CallSite) p10, obj6);
                // ISSUE: reference to a compiler-generated method
                range3.set_Value(System.Type.Missing, (object) arrdblDataTyokyou);
                // ISSUE: variable of a compiler-generated type
                Range usedRange2 = worksheet2.UsedRange;
                // ISSUE: reference to a compiler-generated field
                if (Form1.\u003C\u003Eo__11.\u003C\u003Ep__13 == null)
                {
                  // ISSUE: reference to a compiler-generated field
                  Form1.\u003C\u003Eo__11.\u003C\u003Ep__13 = CallSite<Func<CallSite, object, Range>>.Create(Binder.Convert(CSharpBinderFlags.None, typeof (Range), typeof (Form1)));
                }
                // ISSUE: reference to a compiler-generated field
                Func<CallSite, object, Range> target7 = Form1.\u003C\u003Eo__11.\u003C\u003Ep__13.Target;
                // ISSUE: reference to a compiler-generated field
                CallSite<Func<CallSite, object, Range>> p13 = Form1.\u003C\u003Eo__11.\u003C\u003Ep__13;
                // ISSUE: reference to a compiler-generated field
                if (Form1.\u003C\u003Eo__11.\u003C\u003Ep__12 == null)
                {
                  // ISSUE: reference to a compiler-generated field
                  Form1.\u003C\u003Eo__11.\u003C\u003Ep__12 = CallSite<Func<CallSite, object, object, object, object>>.Create(Binder.GetIndex(CSharpBinderFlags.None, typeof (Form1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[3]
                  {
                    CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
                    CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
                    CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
                  }));
                }
                // ISSUE: reference to a compiler-generated field
                Func<CallSite, object, object, object, object> target8 = Form1.\u003C\u003Eo__11.\u003C\u003Ep__12.Target;
                // ISSUE: reference to a compiler-generated field
                CallSite<Func<CallSite, object, object, object, object>> p12 = Form1.\u003C\u003Eo__11.\u003C\u003Ep__12;
                // ISSUE: reference to a compiler-generated field
                if (Form1.\u003C\u003Eo__11.\u003C\u003Ep__11 == null)
                {
                  // ISSUE: reference to a compiler-generated field
                  Form1.\u003C\u003Eo__11.\u003C\u003Ep__11 = CallSite<Func<CallSite, Worksheet, object>>.Create(Binder.GetMember(CSharpBinderFlags.ResultIndexed, "Range", typeof (Form1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
                  {
                    CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType, (string) null)
                  }));
                }
                // ISSUE: reference to a compiler-generated field
                // ISSUE: reference to a compiler-generated field
                object obj7 = Form1.\u003C\u003Eo__11.\u003C\u003Ep__11.Target((CallSite) Form1.\u003C\u003Eo__11.\u003C\u003Ep__11, worksheet2);
                object cell7 = worksheet2.Cells[(object) 3, (object) 1];
                object cell8 = worksheet2.Cells[(object) usedRange2.Rows.Count, (object) 12];
                object obj8 = target8((CallSite) p12, obj7, cell7, cell8);
                // ISSUE: variable of a compiler-generated type
                Range range4 = target7((CallSite) p13, obj8);
                // ISSUE: reference to a compiler-generated field
                if (Form1.\u003C\u003Eo__11.\u003C\u003Ep__14 == null)
                {
                  // ISSUE: reference to a compiler-generated field
                  Form1.\u003C\u003Eo__11.\u003C\u003Ep__14 = CallSite<Action<CallSite, Range, object, XlSortOrder>>.Create(Binder.InvokeMember(CSharpBinderFlags.ResultDiscarded, "Sort", (IEnumerable<System.Type>) null, typeof (Form1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[3]
                  {
                    CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType, (string) null),
                    CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.NamedArgument, "Key1"),
                    CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant | CSharpArgumentInfoFlags.NamedArgument, "Order1")
                  }));
                }
                // ISSUE: reference to a compiler-generated field
                // ISSUE: reference to a compiler-generated field
                Form1.\u003C\u003Eo__11.\u003C\u003Ep__14.Target((CallSite) Form1.\u003C\u003Eo__11.\u003C\u003Ep__14, range4, worksheet2.Cells[(object) 3, (object) 12], XlSortOrder.xlAscending);
                // ISSUE: reference to a compiler-generated field
                if (Form1.\u003C\u003Eo__11.\u003C\u003Ep__15 == null)
                {
                  // ISSUE: reference to a compiler-generated field
                  Form1.\u003C\u003Eo__11.\u003C\u003Ep__15 = CallSite<Func<CallSite, object, object[,]>>.Create(Binder.Convert(CSharpBinderFlags.None, typeof (object[,]), typeof (Form1)));
                }
                // ISSUE: reference to a compiler-generated field
                // ISSUE: reference to a compiler-generated field
                // ISSUE: reference to a compiler-generated method
                object[,] arrData = Form1.\u003C\u003Eo__11.\u003C\u003Ep__15.Target((CallSite) Form1.\u003C\u003Eo__11.\u003C\u003Ep__15, range4.get_Value(System.Type.Missing));
                operateCsv.ConvertObjectToCsv(arrData, text + str5);
                // ISSUE: reference to a compiler-generated field
                if (Form1.\u003C\u003Eo__11.\u003C\u003Ep__18 == null)
                {
                  // ISSUE: reference to a compiler-generated field
                  Form1.\u003C\u003Eo__11.\u003C\u003Ep__18 = CallSite<Func<CallSite, object, Range>>.Create(Binder.Convert(CSharpBinderFlags.None, typeof (Range), typeof (Form1)));
                }
                // ISSUE: reference to a compiler-generated field
                Func<CallSite, object, Range> target9 = Form1.\u003C\u003Eo__11.\u003C\u003Ep__18.Target;
                // ISSUE: reference to a compiler-generated field
                CallSite<Func<CallSite, object, Range>> p18 = Form1.\u003C\u003Eo__11.\u003C\u003Ep__18;
                // ISSUE: reference to a compiler-generated field
                if (Form1.\u003C\u003Eo__11.\u003C\u003Ep__17 == null)
                {
                  // ISSUE: reference to a compiler-generated field
                  Form1.\u003C\u003Eo__11.\u003C\u003Ep__17 = CallSite<Func<CallSite, object, object, object, object>>.Create(Binder.GetIndex(CSharpBinderFlags.None, typeof (Form1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[3]
                  {
                    CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
                    CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
                    CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
                  }));
                }
                // ISSUE: reference to a compiler-generated field
                Func<CallSite, object, object, object, object> target10 = Form1.\u003C\u003Eo__11.\u003C\u003Ep__17.Target;
                // ISSUE: reference to a compiler-generated field
                CallSite<Func<CallSite, object, object, object, object>> p17 = Form1.\u003C\u003Eo__11.\u003C\u003Ep__17;
                // ISSUE: reference to a compiler-generated field
                if (Form1.\u003C\u003Eo__11.\u003C\u003Ep__16 == null)
                {
                  // ISSUE: reference to a compiler-generated field
                  Form1.\u003C\u003Eo__11.\u003C\u003Ep__16 = CallSite<Func<CallSite, Worksheet, object>>.Create(Binder.GetMember(CSharpBinderFlags.ResultIndexed, "Range", typeof (Form1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
                  {
                    CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType, (string) null)
                  }));
                }
                // ISSUE: reference to a compiler-generated field
                // ISSUE: reference to a compiler-generated field
                object obj9 = Form1.\u003C\u003Eo__11.\u003C\u003Ep__16.Target((CallSite) Form1.\u003C\u003Eo__11.\u003C\u003Ep__16, worksheet2);
                object cell9 = worksheet2.Cells[(object) 2, (object) 152];
                object cell10 = worksheet2.Cells[(object) 22, (object) 158];
                object obj10 = target10((CallSite) p17, obj9, cell9, cell10);
                // ISSUE: variable of a compiler-generated type
                Range range5 = target9((CallSite) p18, obj10);
                // ISSUE: reference to a compiler-generated field
                if (Form1.\u003C\u003Eo__11.\u003C\u003Ep__21 == null)
                {
                  // ISSUE: reference to a compiler-generated field
                  Form1.\u003C\u003Eo__11.\u003C\u003Ep__21 = CallSite<Func<CallSite, object, Range>>.Create(Binder.Convert(CSharpBinderFlags.None, typeof (Range), typeof (Form1)));
                }
                // ISSUE: reference to a compiler-generated field
                Func<CallSite, object, Range> target11 = Form1.\u003C\u003Eo__11.\u003C\u003Ep__21.Target;
                // ISSUE: reference to a compiler-generated field
                CallSite<Func<CallSite, object, Range>> p21 = Form1.\u003C\u003Eo__11.\u003C\u003Ep__21;
                // ISSUE: reference to a compiler-generated field
                if (Form1.\u003C\u003Eo__11.\u003C\u003Ep__20 == null)
                {
                  // ISSUE: reference to a compiler-generated field
                  Form1.\u003C\u003Eo__11.\u003C\u003Ep__20 = CallSite<Func<CallSite, object, object, object, object>>.Create(Binder.GetIndex(CSharpBinderFlags.None, typeof (Form1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[3]
                  {
                    CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
                    CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
                    CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
                  }));
                }
                // ISSUE: reference to a compiler-generated field
                Func<CallSite, object, object, object, object> target12 = Form1.\u003C\u003Eo__11.\u003C\u003Ep__20.Target;
                // ISSUE: reference to a compiler-generated field
                CallSite<Func<CallSite, object, object, object, object>> p20 = Form1.\u003C\u003Eo__11.\u003C\u003Ep__20;
                // ISSUE: reference to a compiler-generated field
                if (Form1.\u003C\u003Eo__11.\u003C\u003Ep__19 == null)
                {
                  // ISSUE: reference to a compiler-generated field
                  Form1.\u003C\u003Eo__11.\u003C\u003Ep__19 = CallSite<Func<CallSite, Worksheet, object>>.Create(Binder.GetMember(CSharpBinderFlags.ResultIndexed, "Range", typeof (Form1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
                  {
                    CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType, (string) null)
                  }));
                }
                // ISSUE: reference to a compiler-generated field
                // ISSUE: reference to a compiler-generated field
                object obj11 = Form1.\u003C\u003Eo__11.\u003C\u003Ep__19.Target((CallSite) Form1.\u003C\u003Eo__11.\u003C\u003Ep__19, worksheet1);
                object cell11 = worksheet1.Cells[(object) (22 * index1 + 1), (object) 1];
                object cell12 = worksheet1.Cells[(object) (22 * index1 + 21), (object) 7];
                object obj12 = target12((CallSite) p20, obj11, cell11, cell12);
                // ISSUE: reference to a compiler-generated method
                // ISSUE: reference to a compiler-generated method
                target11((CallSite) p21, obj12).set_Value(System.Type.Missing, range5.get_Value(System.Type.Missing));
                instance.DisplayAlerts = false;
                // ISSUE: reference to a compiler-generated method
                workbook2.Save();
                instance.DisplayAlerts = true;
                ++index1;
                this.prgDownload.Value = 66 + 34 * index1 / num1;
              }
              ++index2;
            }
            // ISSUE: variable of a compiler-generated type
            Range usedRange = worksheet1.UsedRange;
            // ISSUE: reference to a compiler-generated field
            if (Form1.\u003C\u003Eo__11.\u003C\u003Ep__22 == null)
            {
              // ISSUE: reference to a compiler-generated field
              Form1.\u003C\u003Eo__11.\u003C\u003Ep__22 = CallSite<Func<CallSite, object, object[,]>>.Create(Binder.Convert(CSharpBinderFlags.None, typeof (object[,]), typeof (Form1)));
            }
            // ISSUE: reference to a compiler-generated field
            // ISSUE: reference to a compiler-generated field
            // ISSUE: reference to a compiler-generated method
            object[,] objArray = Form1.\u003C\u003Eo__11.\u003C\u003Ep__22.Target((CallSite) Form1.\u003C\u003Eo__11.\u003C\u003Ep__22, usedRange.get_Value(System.Type.Missing));
            string str6 = "c" + str4 + ".csv";
            instance.DisplayAlerts = false;
            // ISSUE: reference to a compiler-generated method
            worksheet1.SaveAs(text + str6, (object) 6, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing);
            // ISSUE: reference to a compiler-generated method
            workbook1.Close(System.Type.Missing, System.Type.Missing, System.Type.Missing);
            // ISSUE: reference to a compiler-generated method
            workbook2.Close(System.Type.Missing, System.Type.Missing, System.Type.Missing);
            instance.DisplayAlerts = true;
            Marshal.ReleaseComObject((object) worksheet2);
            Marshal.ReleaseComObject((object) workbook2);
            Marshal.ReleaseComObject((object) worksheet1);
            Marshal.ReleaseComObject((object) workbook1);
            Marshal.ReleaseComObject((object) instance);
            this.rtbData.Text = str4 + " 調教データ取得完了しました。";
            this.AxJVLink1.JVClose();
            SystemSounds.Asterisk.Play();
            this.prgDownload.Value = 100;
          }
        }
      }
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

        private void PutTyoukyouDataAllX(
      DateTime datetimeTarg,
      string[] dataTyokyo,
      List<cRaceUma> ListUmas,
      out string[,] arrDataTyokyou,
      out double[,] arrdblDataTyokyou)
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
                            string codeName = this.objCodeConv.GetCodeName("2301", jvHcHanro.TresenKubun, (short)2);
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
    }
}
