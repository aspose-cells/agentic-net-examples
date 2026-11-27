// Title: Auto‑fit Excel worksheet row height to cell content using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# auto fit rows based on cell text | set Excel row height to adjust automatically with Aspose.Cells | C# code to automatically resize worksheet rows to fit content using Aspose.Cells | how to use AutoFitRows method in Aspose.Cells .NET | adjust row height dynamically in generated Excel file with Aspose.Cells
// Tags: Aspose.Cells AutoFitRows method | auto‑fit worksheet rows .NET | dynamic row height Excel C# | Excel row height auto adjustment Aspose.Cells | worksheet row height based on cell content

using Aspose.Cells;

// Creates a new Workbook, adds sample text to cells, calls AutoFitRows() to automatically resize each row's height to fit its content, and saves the file as Result.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        var workbook = new Workbook();

        // Get the first worksheet
        var sheet = workbook.Worksheets[0];

        // (Optional) Add some sample data to demonstrate auto‑fit
        sheet.Cells["A1"].PutValue("Short");
        sheet.Cells["A2"].PutValue("This is a longer piece of text that should cause the row height to increase automatically.");

        // Adjust all rows so their height fits the cell content
        sheet.AutoFitRows();

        // Save the workbook
        workbook.Save("Result.xlsx");
    }
}
