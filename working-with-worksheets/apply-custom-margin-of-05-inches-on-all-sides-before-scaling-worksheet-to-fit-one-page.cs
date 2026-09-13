// Title: Apply 0.5‑inch margins on all sides and fit a worksheet to a single page with Aspose.Cells for .NET (C#)
// AI Prompts: Set the worksheet’s left, right, top, and bottom margins to 0.5 inches (36 points) and configure FitToPagesWide = 1 and FitToPagesTall = 1 using Aspose.Cells in C#. | Create a new workbook, apply half‑inch margins, scale the sheet to one page, add sample cells, and save the file as an XLSX with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# set half‑inch page margins | Fit an Excel worksheet to one page using Aspose.Cells .NET | Convert inches to points for margins in Aspose.Cells C# | Scale worksheet to single page with custom margins Aspose.Cells | PageSetup margin values in points Aspose.Cells C# example
// Tags: set page margins Aspose.Cells C# | worksheet fit-to-page Aspose.Cells | custom margin points Excel .NET | page setup scaling Aspose.Cells | margin conversion inches to points C# | worksheet layout configuration Aspose.Cells

using Aspose.Cells;

// The example creates a new workbook, sets all page margins to 0.5 inches (36 points), configures the sheet to fit on a single page (both width and height), adds a couple of sample cells, and saves the result as Result.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set custom margins (0.5 inches = 36 points)
        sheet.PageSetup.LeftMargin = 36;   // Left margin
        sheet.PageSetup.RightMargin = 36;  // Right margin
        sheet.PageSetup.TopMargin = 36;    // Top margin
        sheet.PageSetup.BottomMargin = 36; // Bottom margin

        // Scale the worksheet to fit one page
        sheet.PageSetup.FitToPagesWide = 1;
        sheet.PageSetup.FitToPagesTall = 1;

        // (Optional) Add some data to demonstrate the layout
        sheet.Cells["A1"].PutValue("Demo");
        sheet.Cells["B2"].PutValue(12345);

        // Save the workbook
        workbook.Save("Result.xlsx");
    }
}
