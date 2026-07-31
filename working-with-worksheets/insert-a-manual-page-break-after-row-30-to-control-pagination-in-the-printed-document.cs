// Title: How to add a manual horizontal page break after row 30 with Aspose.Cells for .NET (C#)
// Description: C# example that creates a workbook with Aspose.Cells, fills the first 50 rows, inserts a manual horizontal page break after row 30 to control printed pagination, and saves the file as PageBreakAfterRow30.xlsx.
// Keywords: Aspose.Cells | C# | .NET | horizontal page break | manual page break | Excel pagination | worksheet page break | programmatic page break | print layout | Aspose.Cells example
// Common Searches: Aspose.Cells add page break after row 30 | C# insert horizontal page break in Excel worksheet | how to control pagination with Aspose.Cells | programmatically set page breaks in .NET Excel file | manual page break Aspose.Cells tutorial
// Developer Intent: Insert a manual horizontal page break after row 30 to manage print pagination in an Excel worksheet using Aspose.Cells.
// Use Cases: Generate printable reports where each section starts on a new page by inserting page breaks at specific rows. | Create multi‑invoice workbooks that print each invoice on a separate page by adding a break after the last row of each invoice. | Prepare large data sheets for printing, adding regular page breaks to keep headers visible on every printed page.
// AI Prompts: Show me how to add horizontal page breaks every 25 rows using Aspose.Cells for .NET. | Provide C# code to insert a vertical page break after column 5 and a horizontal page break after row 30 in the same worksheet. | Explain how to combine manual page breaks with page setup options such as margins, orientation, and scaling in Aspose.Cells.

using Aspose.Cells;

// C# example that creates a workbook with Aspose.Cells, fills the first 50 rows, inserts a manual horizontal page break after row 30 to control printed pagination, and saves the file as PageBreakAfterRow30.xlsx.
class InsertPageBreakDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // (Optional) Fill some data to visualize the page break effect
        for (int i = 0; i < 50; i++)
        {
            worksheet.Cells[i, 0].PutValue($"Row {i + 1}");
        }

        // Insert a manual horizontal page break after row 30 (zero‑based index)
        worksheet.HorizontalPageBreaks.Add(30);

        // Save the workbook with the page break applied
        workbook.Save("PageBreakAfterRow30.xlsx");
    }
}
