using System.IO;
using System.Text;

namespace getJVLInkData
{
    internal class OperateCSV
    {
        public void ConvertObjectToCsv(object[,] arrData, string csvPath)
        {
            Encoding encoding = Encoding.GetEncoding("Shift_JIS");
            StreamWriter streamWriter = new StreamWriter(csvPath, false, encoding);
            int length1 = arrData.GetLength(0);
            int length2 = arrData.GetLength(1);
            int num = length2;
            for (int index1 = 1; index1 <= length1 && arrData[index1, 1] != null; ++index1)
            {
                for (int index2 = 1; index2 <= length2; ++index2)
                {
                    string field = "";
                    if (arrData[index1, index2] != null)
                        field = arrData[index1, index2].ToString();
                    string s = this.EncloseDoubleQuotesIfNeed(field);
                    int result = 0;
                    if (int.TryParse(s, out result))
                        streamWriter.Write(result);
                    else
                        streamWriter.Write(s);
                    if (num > index2)
                        streamWriter.Write(',');
                }
                if (arrData[index1, 1] != null)
                    streamWriter.Write("\r\n");
            }
            streamWriter.Close();
        }

        private string EncloseDoubleQuotesIfNeed(string field)
        {
            return this.NeedEncloseDoubleQuotes(field) ? this.EncloseDoubleQuotes(field) : field;
        }

        private string EncloseDoubleQuotes(string field)
        {
            if (field.IndexOf('"') > -1)
                field = field.Replace("\"", "\"\"");
            return "\"" + field + "\"";
        }

        private bool NeedEncloseDoubleQuotes(string field)
        {
            return field.IndexOf('"') > -1 || field.IndexOf(',') > -1 || (field.IndexOf('\r') > -1 || field.IndexOf('\n') > -1) || (field.StartsWith(" ") || field.StartsWith("\t") || field.EndsWith(" ")) || field.EndsWith("\t");
        }
    }
}
