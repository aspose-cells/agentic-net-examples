using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a hyperlink to cell A1 with display text and a screen tip
        // Parameters: startCell, endCell, address, text to display, screen tip
        worksheet.Hyperlinks.Add("A1", "A1", "https://www.aspose.com", "Visit Aspose", "Open Aspose website");

        // Save the workbook
        workbook.Save("HyperlinkDemo.xlsx");
    }
}