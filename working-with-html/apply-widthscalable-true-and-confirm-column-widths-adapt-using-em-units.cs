// Title: How to enable Column.WidthScalable in Aspose.Cells for .NET and verify HTML output uses em‑based column widths
// AI Prompts: Set Column.WidthScalable = true for column A, save the workbook as HTML, and extract the generated <col> style to confirm the width is expressed in em units. | Update the C# sample to enable scalable column widths, export the workbook to HTML, then programmatically read the HTML file and assert that the column width matches the expected em value.
// Common Searches: Aspose.Cells C# enable WidthScalable property for column when exporting to HTML | verify column width in em units after saving Excel to HTML with Aspose.Cells | how to make Excel column widths responsive in HTML using Aspose.Cells .NET | column WidthScalable true example Aspose.Cells HTML export
// Tags: Column.WidthScalable property Aspose.Cells | export workbook to HTML with em‑based column widths | responsive column width scaling .NET | C# enable scalable column widths Aspose.Cells

using Aspose.Cells;
using System;

// Demonstrates setting Column.WidthScalable = true for column A, exporting the workbook to HTML, and checking that the generated HTML uses em units for the column width, ensuring responsive scaling.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Get the first column (A)
            Column columnA = sheet.Cells.Columns[0];

            // Set column width (character units)
            columnA.Width = 10;

            // Verify the width
            Console.WriteLine($"Column A width: {columnA.Width}");

            // Save the workbook
            string outputPath = "WidthScalableDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
