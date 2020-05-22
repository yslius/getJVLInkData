using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using System.Drawing.Imaging;

namespace getJVLInkData
{
    public class clsTrain
    {
        Form1 _form1;
        string[] arrAddHead = { "単勝配当", "1着複勝配当", "2着複勝配当",
            "3着複勝配当", "枠連配当", "馬連配当",
            "馬単配当", "3連複配当", "3連単配当",
            "ワイド1", "ワイド1配当", "ワイド2",
            "ワイド2配当", "ワイド3", "ワイド3配当" };
        string[] arrAddTrain = { "検出数", "調教状態", "補足", "調教指数", "暗号データ", "暗号データ" };

        public clsTrain(Form1 form1)
        {
            _form1 = form1;
        }
        public void ReflectTrainMain()
        {
            //Console.WriteLine(_form1.textBox2.Text);
            string pathTarg;
            string pathFileT;
            string pathFileR;
            List<string> listTcsv;
            List<string> listRcsv;

            pathTarg = _form1.textBox2.Text;

            // 調教データの読み込み
            pathFileT = getTrainDataFile(pathTarg);
            if (pathFileT == "")
                return;

            // 出馬表の読み込み
            pathFileR = GetRaceCardFile(pathTarg);
            if (pathFileR == "")
                return;

            // 調教データの読み込み
            listTcsv = ReadCSV(pathFileT);

            // 出馬表の読み込み
            listRcsv = ReadCSV(pathFileR);

            // 追加項目を記入
            listRcsv = WriteAddHeadData(listRcsv);

            // 追加項目を記入
            //listRcsv = CopyTraindataToRacecard(listRcsv, listRcsv);

            // ファイル出力
            //OutputCsv(listRcsv);

        }

        string getTrainDataFile(string pathTarg)
        {
            string[] listfiles = Directory.GetFiles(pathTarg, "*", SearchOption.AllDirectories);

            foreach(string strfile in listfiles)
            {
                if (System.Text.RegularExpressions.Regex.IsMatch(strfile, @"c\d{8}.csv"))
                {
                    Console.WriteLine(strfile);
                }
            }
            

            string path = pathTarg + "01出馬表.csv";
            if (!File.Exists(path))
            {
                MessageBox.Show("出馬表が見つかりません。", "エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Hand);
                return "";
            }
            return path;
        }

        void checkDateTrainData(string pathfileT, string pathTarg)
        {
            string strDateFolder;
            string strDateFile;
            string strTmp;

            strTmp = pathfileT;



        }

        string GetRaceCardFile(string pathTarg)
        {
            string path = pathTarg + "01出馬表.csv";
            if (!File.Exists(path))
            {
                MessageBox.Show("出馬表が見つかりません。", "エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Hand);
                return "";
            }
            return path;
        }

        List<string> ReadCSV(string pathTarg)
        {
            List<string> lists = new List<string>();
            StreamReader sr = new StreamReader(pathTarg);
            {
                while (!sr.EndOfStream)
                {
                    string line = sr.ReadLine();
                    string[] values = line.Split(',');
                    lists.AddRange(values);
                }
            }
            return lists;
        }

        List<string> WriteAddHeadData(List<string> listcsv)
        {
            foreach(string list in listcsv)
            {

            }
            return listcsv;
        }
    }
}
