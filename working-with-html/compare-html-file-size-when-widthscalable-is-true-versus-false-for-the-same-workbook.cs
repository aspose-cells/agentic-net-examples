// Title: Aspose.Cells C# – Compare HTML Output Size with WidthScalable True vs False
// Description: Creates a workbook with 100 rows × 10 columns, saves two HTML files—one using HtmlSaveOptions.WidthScalable = true (scalable units) and another with WidthScalable = false (pixel units)—then reads the file sizes and reports which setting yields the smaller HTML file.
// Keywords: Aspose.Cells | C# HTML export | HtmlSaveOptions WidthScalable | scalable vs fixed column width | HTML file size comparison | Aspose.Cells performance | Excel to HTML conversion
// Common Searches: Aspose.Cells WidthScalable effect on HTML size | compare scalable and fixed column width in Aspose.Cells | HTML export file size Aspose.Cells C# | which WidthScalable setting creates smaller HTML | measure HTML output size Aspose.Cells
// Developer Intent: Find out whether setting HtmlSaveOptions.WidthScalable to true or false produces a smaller HTML file for the same workbook.
// Use Cases: Select the optimal WidthScalable value to minimize HTML payload for web‑based spreadsheet viewers. | Automate batch conversion of Excel workbooks to HTML and enforce a file‑size limit. | Generate lightweight HTML reports by testing scalable versus fixed column‑width modes.
// AI Prompts: Generate C# code that logs the exact byte difference between the scalable and fixed HTML files produced by Aspose.Cells. | Explain how WidthScalable changes the generated CSS/HTML markup and why it can reduce file size. | Provide a strategy to programmatically choose WidthScalable = true or false based on a maximum allowed HTML size.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsWidthScalableComparison
{
    // Creates a workbook with 100 rows × 10 columns, saves two HTML files—one using HtmlSaveOptions.WidthScalable = true (scalable units) and another with WidthScalable = false (pixel units)—then reads the file sizes and reports which setting yields the smaller HTML file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with enough data to make column widths noticeable
            for (int row = 0; row < 100; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1} - Some long text to test width");
                }
            }

            // Define file names for the two HTML outputs
            string scalablePath = "output_scalable.html";
            string fixedPath = "output_fixed.html";

            // Save with WidthScalable = true
            HtmlSaveOptions scalableOptions = new HtmlSaveOptions();
            scalableOptions.WidthScalable = true; // Export column width using scalable units
            workbook.Save(scalablePath, scalableOptions);

            // Save with WidthScalable = false
            HtmlSaveOptions fixedOptions = new HtmlSaveOptions();
            fixedOptions.WidthScalable = false; // Export column width using fixed pixel units
            workbook.Save(fixedPath, fixedOptions);

            // Get file sizes
            long scalableSize = new FileInfo(scalablePath).Length;
            long fixedSize = new FileInfo(fixedPath).Length;

            // Output the comparison
            Console.WriteLine($"HTML size with WidthScalable = true : {scalableSize} bytes");
            Console.WriteLine($"HTML size with WidthScalable = false: {fixedSize} bytes");

            if (scalableSize < fixedSize)
                Console.WriteLine("Scalable width produces a smaller HTML file.");
            else if (scalableSize > fixedSize)
                Console.WriteLine("Fixed width produces a smaller HTML file.");
            else
                Console.WriteLine("Both files have the same size.");
        }
    }
}
