[ComVisible(true)]
public class WinForm : Form
{
    XmlDocument xml = new XmlDocument();

    public WinForm()
    {
        XmlDeclaration dec = xml.CreateXmlDeclaration("1.0", "utf-8", null);
        xml.AppendChild(dec);

        XmlElement xml_root = xml.CreateElement("root");
        xml.AppendChild(xml_root);
    }

    public string GetStructure()
    {
        XmlDocument xml = new XmlDocument();

        XmlDeclaration dec = xml.CreateXmlDeclaration("1.0", "utf-8", null);
        xml.AppendChild(dec);

        XmlElement xml_root = xml.CreateElement("root");
        xml.AppendChild(xml_root);

        int xml_current_level = 1;

        try
        {
            Document doc = AutoWordAddIn.Plugin.currentDoc;
            foreach (Paragraph item in doc.Paragraphs)
            {
                if (item.OutlineLevel != WdOutlineLevel.wdOutlineLevelBodyText)
                {
                    XmlElement xml_item = xml.CreateElement("item");
                    xml_item.SetAttribute("name", item.Range.Text.Substring(0, item.Range.Text.Length - 1).Trim().Replace("/", "\\/"));
                    xml_item.SetAttribute("level", ((int)item.OutlineLevel).ToString());
                    xml_item.SetAttribute("srange", item.Range.Start.ToString());
                    xml_item.SetAttribute("erange", item.Range.End.ToString());

                    string xml_current_path = "(//item)[last()]";
                    if ((int)item.OutlineLevel == 1)
                    {
                        xml_current_path = "/root";
                    }
                    else
                    {
                        if ((int)item.OutlineLevel == xml_current_level)
                        {
                            xml_current_path += "/..";
                        }
                        else if ((int)item.OutlineLevel < xml_current_level)
                        {
                            for (int i = 0; i < xml_current_level - (int)item.OutlineLevel + 1; i++)
                            {
                                xml_current_path += "/..";
                            }
                        }
                    }

                    XmlNode parentNode = xml.SelectSingleNode(xml_current_path);
                    string parentNames = "";
                    while (parentNode != null && parentNode.Attributes.Count > 0)
                    {
                        parentNames = parentNode.Attributes["name"].Value + "/" + parentNames;
                        parentNode = parentNode.ParentNode;
                    }
                    xml_item.SetAttribute("names", parentNames + xml_item.Attributes["name"].Value);

                    parentNode = xml.SelectSingleNode(xml_current_path);
                    parentNode.AppendChild(xml_item);

                    xml_current_level = (int)item.OutlineLevel;
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message);
        }

        MemoryStream memoryStream = new MemoryStream();
        XmlTextWriter writer = new XmlTextWriter(memoryStream, null)
        {
            Formatting = System.Xml.Formatting.Indented
        };
        xml.Save(writer);

        StreamReader streamReader = new StreamReader(memoryStream);
        memoryStream.Position = 0;
        string xmlString = streamReader.ReadToEnd();
        streamReader.Close();
        memoryStream.Close();

        return xmlString;
    }

    private static string removeEnd(string str)
    {
        return str.Substring(0, str.Length - 1).Trim();
    }

    private static void RecognizeXML(Document doc, XmlDocument xml)
    {
        xml.SelectSingleNode("/root").RemoveAll();

        int xml_current_level = 1;
        try
        {
            foreach (Paragraph item in doc.Paragraphs)
            {
                if (item.OutlineLevel != WdOutlineLevel.wdOutlineLevelBodyText)
                {
                    XmlElement xml_item = xml.CreateElement("item");
                    xml_item.SetAttribute("name", removeEnd(item.Range.Text));
                    xml_item.SetAttribute("level", ((int)item.OutlineLevel).ToString());
                    xml_item.SetAttribute("srange", item.Range.Start.ToString());
                    xml_item.SetAttribute("erange", item.Range.End.ToString());

                    string xml_current_path = "(//item)[last()]";
                    if ((int)item.OutlineLevel == 1)
                    {
                        xml_current_path = "/root";
                    }
                    else
                    {
                        if ((int)item.OutlineLevel == xml_current_level)
                        {
                            xml_current_path += "/..";
                        }
                        else if ((int)item.OutlineLevel < xml_current_level)
                        {
                            for (int i = 0; i < xml_current_level - (int)item.OutlineLevel + 1; i++)
                            {
                                xml_current_path += "/..";
                            }
                        }
                    }

                    XmlNode parentNode = xml.SelectSingleNode(xml_current_path);
                    parentNode.AppendChild(xml_item);

                    xml_current_level = (int)item.OutlineLevel;
                }
            }
        }
        catch (Exception)
        {
        }
    }

    public string Execute(string json)
    {
        try
        {
            Dictionary<string, object> dic = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);

            int index = dic.ContainsKey("index") ? Convert.ToInt32(dic["index"]) : 0;
            string startCatalog = (dic.ContainsKey("input_1") ? dic["input_1"].ToString() : "").Replace("\\/", "```");
            int tableIndex = dic.ContainsKey("input_2") ? Convert.ToInt32(dic["input_2"].ToString()) : 1;
            int rowIndex = dic.ContainsKey("input_3") ? Convert.ToInt32(dic["input_3"].ToString()) : 1;
            int columnIndex = dic.ContainsKey("input_4") ? Convert.ToInt32(dic["input_4"].ToString()) : 1;

            object missing = Type.Missing;

            Document document = AutoWordAddIn.Plugin.currentDoc;

            if (index == 0)
            {
                RecognizeXML(document, xml);
            }

            object pos = 0;
            if (!string.IsNullOrEmpty(startCatalog) && startCatalog.Trim(new char[] { '/' }).Length > 0)
            {
                string[] cs = Regex.Split(startCatalog.Trim(new char[] { '/' }), "/");
                string xpath = "/root";
                foreach (var s in cs)
                {
                    xpath += "/item[@name='" + s.Replace("```", "\\/") + "']";
                }
                XmlNode node = xml.SelectSingleNode(xpath);
                if (node != null)
                {
                    pos = Convert.ToInt32(node.Attributes["erange"].Value);
                }
            }

            Table table = document.Range(ref pos, ref missing).Tables[tableIndex];

            string result = table.Cell(rowIndex, columnIndex).Range.Text;
            result = removeEnd(result);

            return result;
        }
        catch (Exception e)
        {
            return "Exception:" + e.Message;
        }
    }

    public string GetData()
    {
        string dataPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        dataPath += "\\BlueAmber\\AutoWordAddIn";

        string dataFile = dataPath + "\\data-extract.json";

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

        string dataFile = dataPath + "\\data-extract.json";

        System.IO.File.WriteAllText(dataFile, json);

        return "已保存至 " + dataFile;
    }

    public string SetExcel(string json)
    {
        try
        {
            List<Dictionary<string, string>> dics = JsonConvert.DeserializeObject<List<Dictionary<string, string>>>(json);

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = true;
            Microsoft.Office.Interop.Excel.Workbook wb = excel.Workbooks.Add();
            Microsoft.Office.Interop.Excel.Worksheet ws = wb.Sheets[1] as Microsoft.Office.Interop.Excel.Worksheet;

            for (int i = 0; i < dics.Count; i++)
            {
                ws.Cells[i + 1, 1] = dics[i].ContainsKey("input_5") ? dics[i]["input_5"] : "";
                ws.Cells[i + 1, 2] = dics[i].ContainsKey("input_6") ? dics[i]["input_6"] : "";
            }

            return dics.Count.ToString();
        }
        catch (Exception e)
        {
            return "Exception:" + e.Message;
        }
    }
}
