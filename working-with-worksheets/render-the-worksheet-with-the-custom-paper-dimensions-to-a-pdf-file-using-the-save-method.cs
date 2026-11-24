// Title: Render an Excel worksheet with a specific paper size to PDF using Aspose.Cells Save method in C#
// Common Searches: how to set worksheet page size to Letter and export to PDF with Aspose.Cells C# | Aspose.Cells C# save workbook as PDF with custom page setup | C# generate PDF from Excel sheet with specific paper dimensions using Aspose.Cells | using Aspose.Cells to render Excel worksheet to PDF with defined paper size | Aspose.Cells SaveFormat.Pdf with page setup configuration in .NET
// Tags: Aspose.Cells set worksheet paper size | Aspose.Cells export worksheet to PDF | C# Aspose.Cells SaveFormat.Pdf usage | page setup paper size Letter Aspose.Cells | render Excel to PDF with custom dimensions .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing Excel file or creates a new workbook, accesses the first worksheet, sets its page setup to the Letter paper size, ensures the output directory exists, and then saves the workbook as a PDF using the Workbook.Save method with SaveFormat.Pdf, handling any exceptions that may occur.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                Workbook workbook;

                // Load workbook if file exists; otherwise create a new one
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook(); // creates a default workbook with one worksheet
                }

                Worksheet worksheet = workbook.Worksheets[0];

                // Set paper size to Letter (8.5" x 11") – standard size, no need for custom dimensions
                worksheet.PageSetup.PaperSize = PaperSizeType.PaperLetter;

                // Save the workbook as PDF
                string outputPath = "output.pdf";

                // Ensure the directory for the output file exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath, SaveFormat.Pdf);
                Console.WriteLine($"PDF saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
