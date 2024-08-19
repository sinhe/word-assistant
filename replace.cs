[ComVisible(true)]
public class WinForm : Form
{
    public int Execute(string json)
    {
        int founds = 0;
        Dictionary<string, string> dic = JsonConvert.DeserializeObject<Dictionary<string, string>>(json);

        string input_1 = dic.ContainsKey("input_1") ? dic["input_1"] : "";
        string input_2 = dic.ContainsKey("input_2") ? dic["input_2"] : "";
        object input_3 = dic.ContainsKey("input_3") ? dic["input_3"] : "";

        object missing = Type.Missing;

        Document document = AutoWordAddIn.Plugin.currentDoc;
        Microsoft.Office.Interop.Word.Range rng = document.Content;

        rng.Find.ClearFormatting();
        rng.Find.Forward = true;
        rng.Find.Text = input_1;
        rng.Find.Execute(
            ref missing, ref missing, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing);
        while (rng.Find.Found)
        {
            founds++;
            
            if (input_3.ToString().Length > 0)
            {
                document.Comments.Add(document.Range(rng.Start, rng.End), ref input_3);
            }

            rng.Find.Execute(
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing);
        }

        rng = document.Content;
        rng.Find.ClearFormatting();
        rng.Find.Forward = true;
        rng.Find.Text = input_1;
        rng.Find.Replacement.ClearFormatting();
        rng.Find.Replacement.Text = input_2;

        object replaceAll = WdReplace.wdReplaceAll;
        rng.Find.Execute(ref missing, ref missing, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing,
            ref replaceAll, ref missing, ref missing, ref missing, ref missing);

        return founds;
    }

    public string GetData()
    {
        string dataPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        dataPath += "\\BlueAmber\\AutoWordAddIn";

        string dataFile = dataPath + "\\data-replace.json";
    
        using (System.IO.FileStream fs = new System.IO.FileStream(dataFile, System.IO.FileMode.OpenOrCreate, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite))
        {
            byte[] fsb = new byte[fs.Length];
            fs.Read(fsb, 0, fsb.Length);
            return Encoding.UTF8.GetString(fsb);
        }
    }

    public string SetData(string json)
    {
        string dataPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        dataPath += "\\BlueAmber\\AutoWordAddIn";

        string dataFile = dataPath + "\\data-replace.json";
        
        System.IO.File.WriteAllText(dataFile, json);

        return "已保存至 " + dataFile;
    }

    public string GetExcel()
    {
        OpenFileDialog fileDialog = new OpenFileDialog();
        fileDialog.Filter = "Excel 文件|*.xls;*.xlsx";
        if (fileDialog.ShowDialog() == DialogResult.OK)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = false;
            Microsoft.Office.Interop.Excel.Workbook wb = excel.Workbooks.Open(fileDialog.FileName);
            Microsoft.Office.Interop.Excel.Worksheet ws = wb.Sheets[1] as Microsoft.Office.Interop.Excel.Worksheet;
            Microsoft.Office.Interop.Excel.Range rg = ws.UsedRange;

            List<Dictionary<string, object>> dics = new List<Dictionary<string, object>>();
            for (int i = 1; i <= rg.Rows.Count; i++)
            {
                Dictionary<string, object> dic = new Dictionary<string, object>();
                for (int j = 1; j <= rg.Columns.Count; j++)
                {
                    Microsoft.Office.Interop.Excel.Range cell = rg.Cells[i, j] as Microsoft.Office.Interop.Excel.Range;
                    dic["input_" + j.ToString()] = cell.Text;
                }
                dics.Add(dic);
            }

            wb.Close();
            excel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);

            return JsonConvert.SerializeObject(dics);
        }

        return "[]";
    }
}
