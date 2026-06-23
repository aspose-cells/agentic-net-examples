using Aspose.Cells;

class InsertPageBreakDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some rows (optional, just for demonstration)
        for (int i = 0; i < 50; i++)
        {
            worksheet.Cells[i, 0].PutValue($"Row {i + 1}");
        }

        // Insert a manual horizontal page break after row 30 (zero‑based index)
        worksheet.HorizontalPageBreaks.Add(30);

        // Save the workbook
        workbook.Save("PageBreakAfterRow30.xlsx");
    }
}