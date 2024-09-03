[ComVisible(true)]
public class WinForm : Form
{
    public string Execute()
    {
        Document document = AutoWordAddIn.Plugin.currentDoc;

        List<string> styles = new List<string>();
        List<string> fonts = new List<string>();
        List<float> fontSizes = new List<float>();

        foreach (Paragraph paragraph in document.Paragraphs)
        {
            Microsoft.Office.Interop.Word.Style paragraphStyle = paragraph.get_Style() as Microsoft.Office.Interop.Word.Style;
            if (paragraphStyle != null && !styles.Contains(paragraphStyle.NameLocal))
            {
                styles.Add(paragraphStyle.NameLocal);
            }

            string fontName = paragraph.Range.Font.Name;
            float fontSize = paragraph.Range.Font.Size;

            if (!fonts.Contains(fontName))
            {
                fonts.Add(fontName);
            }
            if (!fontSizes.Contains(fontSize))
            {
                fontSizes.Add(fontSize);
            }
        }

        foreach (Table table in document.Tables)
        {
            for (int i = 1; i <= table.Rows.Count; i++)
            {
                try
                {
                    Row row = table.Rows[i];
                    for (int j = 1; j < row.Cells.Count; j++)
                    {
                        try
                        {
                            Cell cell = row.Cells[j];

                            string cellFontName = cell.Range.Font.Name;
                            float cellFontSize = cell.Range.Font.Size;

                            if (!fonts.Contains(cellFontName))
                            {
                                fonts.Add(cellFontName);
                            }
                            if (!fontSizes.Contains(cellFontSize))
                            {
                                fontSizes.Add(cellFontSize);
                            }
                        }
                        catch (Exception)
                        { }
                    }
                }
                catch (Exception)
                { }
            }
        }

        return JsonConvert.SerializeObject(new
        {
            styles = styles,
            fonts = fonts,
            fontSizes = fontSizes
        });
    }
}
