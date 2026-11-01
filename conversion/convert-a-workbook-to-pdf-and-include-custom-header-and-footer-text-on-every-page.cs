// Title: Convert an Excel workbook to PDF with custom header and footer using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add left header and center footer before saving workbook as PDF | How to set page header and footer for all worksheets in Aspose.Cells and export to PDF | C# export Excel to PDF with dynamic date footer using Aspose.Cells | Apply company confidential header to every page of PDF generated from Excel with Aspose.Cells | Aspose.Cells PDF conversion include page numbers and custom footer text
// Tags: Aspose.Cells set worksheet header C# | Aspose.Cells add footer to PDF export | Excel to PDF conversion with custom page header | C# Aspose.Cells page setup header footer | PDF generation from workbook with header footer Aspose

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file, iterates through each worksheet to assign a left header and a centered footer via the PageSetup object, ensures the output folder exists, and then saves the workbook as a PDF that contains the custom header and footer on every page.
class WorkbookToPdfWithHeaderFooter
{
    static void Main()
    {
        try
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Define custom header and footer text
            string headerText = "Company Confidential - Report";
            string footerText = "Generated on " + DateTime.Now.ToString("yyyy-MM-dd") + " - Page &P of &N";

            // Apply header and footer to every worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the page setup of the worksheet
                PageSetup pageSetup = sheet.PageSetup;

                // Set the left header (section 0) and center footer (section 1)
                pageSetup.SetHeader(0, headerText);   // Left header
                pageSetup.SetFooter(1, footerText);   // Center footer

                // Optional: ensure headings are printed (if required)
                pageSetup.PrintHeadings = true;
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as PDF with the applied header/footer
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
