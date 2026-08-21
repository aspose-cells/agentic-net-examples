// Title: Aspose.Cells for .NET – Export Excel to HTML with scalable column widths (em units)
// Description: C# example that creates a workbook, sets column widths, fills cells, and uses HtmlSaveOptions.WidthScalable = true to generate HTML where column widths are expressed in em units. The demo also saves a second file with WidthScalable = false for side‑by‑side comparison, illustrating responsive versus fixed column sizing.
// Keywords: Aspose.Cells | C# | HtmlSaveOptions | WidthScalable | scalable column width | em units | Excel to HTML export | responsive HTML report | column width scaling | Aspose.Cells example
// Common Searches: Aspose.Cells export Excel to HTML with em units | HtmlSaveOptions WidthScalable true example | How to make HTML column widths responsive in Aspose.Cells | Difference between WidthScalable true and false | C# code for scalable column widths in HTML output
// Developer Intent: Generate HTML from an Excel workbook where column widths automatically adapt to the surrounding font size by enabling WidthScalable.
// Use Cases: Create responsive web reports that adjust column widths with user‑defined font sizes. | Produce two versions of an HTML export—one scalable, one fixed—to test layout behavior across devices. | Build email‑ready HTML tables whose column dimensions scale with different email client settings. | Integrate scalable HTML export into a .NET web application that serves dynamic spreadsheet data.
// AI Prompts: Write a script that parses the saved HTML file and verifies that column widths are defined using ‘em’ units after setting WidthScalable to true. | Explain how changing the base font size in the generated HTML affects column width scaling when WidthScalable is enabled. | Provide a step‑by‑step guide to compare the visual differences between the scalable and fixed HTML files produced by the example.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // C# example that creates a workbook, sets column widths, fills cells, and uses HtmlSaveOptions.WidthScalable = true to generate HTML where column widths are expressed in em units. The demo also saves a second file with WidthScalable = false for side‑by‑side comparison, illustrating responsive versus fixed column sizing.
    public class WidthScalableDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Set column widths (in character units) for demonstration
                cells.SetColumnWidth(0, 20); // Column A
                cells.SetColumnWidth(1, 30); // Column B

                // Populate cells with sample data
                cells["A1"].PutValue("Short");
                cells["B1"].PutValue("This is a longer text that will require more width");

                // Configure HTML save options to export column widths using scalable units (em)
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
                htmlOptions.WidthScalable = true; // Enable scalable column width

                // Save the workbook as HTML with scalable column widths
                workbook.Save("output_scalable.html", htmlOptions);

                // Save again with fixed column widths for comparison
                htmlOptions.WidthScalable = false;
                workbook.Save("output_fixed.html", htmlOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            WidthScalableDemo.Run();
        }
    }
}
