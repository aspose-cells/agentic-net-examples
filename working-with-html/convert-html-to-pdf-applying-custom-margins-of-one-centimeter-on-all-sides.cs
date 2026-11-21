// Title: Convert HTML to PDF with 1 cm page margins using Aspose.Cells in C#
// Common Searches: c# set 1 cm page margins when converting HTML to PDF with Aspose.Cells | Aspose.Cells HTML to PDF conversion with custom margins example | how to configure PageSetup margins in centimeters in Aspose.Cells | convert HTML tables to PDF and set uniform margins using Aspose.Cells .NET | sample code for HTML to PDF conversion with margin settings in Aspose.Cells
// Tags: Aspose.Cells HTML to PDF with centimeter margins | PageSetup margin configuration Aspose.Cells C# | Workbook.Save PDF with custom margins Aspose.Cells | convert HTML tables to PDF using Aspose.Cells | set uniform page margins in Aspose.Cells workbook

using System;
using System.IO;
using Aspose.Cells;

// The example verifies the input HTML file, loads it into an Aspose.Cells Workbook, accesses the first worksheet, sets all page margins to 1 cm (converted to inches) via PageSetup, and saves the workbook as a PDF, handling any runtime exceptions.
class HtmlToPdfConverter
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.html";
            const string outputPath = "output.pdf";

            // Verify that the input HTML file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the HTML file into a workbook. Aspose.Cells can interpret HTML tables.
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (or iterate if needed).
            Worksheet sheet = workbook.Worksheets[0];

            // Set margins to 1 centimeter (≈0.3937 inches) on all sides.
            // PageSetup margins are measured in inches in recent Aspose.Cells versions.
            const double cmToInch = 0.393701; // 1 cm in inches
            PageSetup pageSetup = sheet.PageSetup;
            pageSetup.LeftMargin = cmToInch;
            pageSetup.RightMargin = cmToInch;
            pageSetup.TopMargin = cmToInch;
            pageSetup.BottomMargin = cmToInch;

            // Save the workbook as a PDF file.
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF successfully created at '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
