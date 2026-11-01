// Title: Convert an HTML file to PDF with 1 cm margins using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an HTML document into an Aspose.Cells Workbook, sets 1 cm top, bottom, left, and right margins, and saves the workbook as a PDF. | Demonstrate how to configure page margins in centimeters for a worksheet before exporting to PDF with Aspose.Cells. | Show the steps to verify the input HTML file exists and handle errors while converting to PDF with custom margins in C#.
// Common Searches: Aspose.Cells C# convert html to pdf with 1 cm margins | how to set worksheet margins in centimeters before PDF export using Aspose.Cells | C# example loading HTML into workbook and exporting to PDF with uniform margins | custom page margins for PDF output from HTML in Aspose.Cells .NET | error handling when converting HTML to PDF with Aspose.Cells
// Tags: Aspose.Cells HTML to PDF with custom margins | set worksheet margins centimeters Aspose.Cells | C# convert HTML workbook to PDF | page setup margins Aspose.Cells API | load HTML into Aspose.Cells workbook | export workbook as PDF .NET

using System;
using System.IO;
using Aspose.Cells;

// Loads an HTML file into an Aspose.Cells Workbook, applies 1 cm margins on all sides of the first worksheet, and saves the result as a PDF.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.html";
            string outputPath = "output.pdf";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load HTML document into a workbook
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Convert 1 centimeter to points (1 cm = 28.3465 points)
            const double cmToPoints = 28.3465;

            // Apply 1 cm margins on all sides to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            PageSetup pageSetup = sheet.PageSetup;
            pageSetup.TopMargin = cmToPoints;
            pageSetup.BottomMargin = cmToPoints;
            pageSetup.LeftMargin = cmToPoints;
            pageSetup.RightMargin = cmToPoints;

            // Save the workbook as PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
