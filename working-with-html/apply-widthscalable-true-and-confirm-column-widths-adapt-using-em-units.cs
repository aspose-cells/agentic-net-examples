// Title: Export Excel to HTML with Scalable Column Widths (WidthScalable = true) in C#
// Description: Shows how to enable HtmlSaveOptions.WidthScalable in Aspose.Cells so column widths are written in em units, then saves a second file with fixed pixel widths for side‑by‑side comparison.
// Keywords: Aspose.Cells | HtmlSaveOptions | WidthScalable | C# | column width em units | responsive HTML export | Excel to HTML | scalable columns | .NET | HTMLSaveOptions example
// Common Searches: Aspose.Cells WidthScalable example C# | HTML export column widths in em units | How to make Excel HTML responsive with Aspose | compare scalable vs fixed column widths Aspose.Cells | C# save workbook as HTML with em column sizes
// Developer Intent: Turn on WidthScalable to generate HTML where column widths adapt to font size using em units.
// Use Cases: Create responsive HTML reports that adjust to different screen resolutions. | Produce two HTML versions—one scalable, one fixed—to test layout behavior. | Embed Excel data in email templates where column sizing must follow the surrounding text size.
// AI Prompts: Write a script that parses output_scalable.html and verifies that each column style uses an 'em' value. | Explain how to calculate appropriate column widths before enabling WidthScalable for a desired layout. | Provide step‑by‑step instructions to visually compare output_scalable.html with output_fixed.html in a browser.

using System;
using Aspose.Cells;

namespace AsposeCellsWidthScalableDemo
{
    // Shows how to enable HtmlSaveOptions.WidthScalable in Aspose.Cells so column widths are written in em units, then saves a second file with fixed pixel widths for side‑by‑side comparison.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some data to make column widths visible
            cells["A1"].PutValue("Short");
            cells["B1"].PutValue("This is a considerably longer piece of text");
            cells["C1"].PutValue("Medium length");

            // Set explicit column widths (in character units)
            // These widths will be converted to scalable units (em) when WidthScalable is true
            cells.SetColumnWidth(0, 12); // Column A
            cells.SetColumnWidth(1, 30); // Column B
            cells.SetColumnWidth(2, 20); // Column C

            // Configure HTML save options to use scalable column widths
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.WidthScalable = true;   // Enable scalable widths (em units)
            htmlOptions.ImageScalable = true;   // Keep default for images

            // Save the workbook as HTML with scalable column widths
            workbook.Save("output_scalable.html", htmlOptions);

            // Save again with fixed widths for comparison
            htmlOptions.WidthScalable = false;
            workbook.Save("output_fixed.html", htmlOptions);

            // Inform the user that the files have been created
            Console.WriteLine("HTML files generated:");
            Console.WriteLine(" - output_scalable.html (column widths in em units)");
            Console.WriteLine(" - output_fixed.html (column widths in pixels)");
        }
    }
}
