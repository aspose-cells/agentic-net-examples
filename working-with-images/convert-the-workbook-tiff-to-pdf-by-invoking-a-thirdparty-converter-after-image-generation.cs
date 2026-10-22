// Title: Convert an Excel .xlsx workbook to PDF in a C# console app using Aspose.Cells
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, verifies the file exists, creates the output folder if missing, and saves the workbook as a PDF with comprehensive exception handling. | Write a C# console program that demonstrates converting a spreadsheet to PDF using Aspose.Cells, including input‑file validation and automatic creation of the destination directory. | Create a C# snippet that shows how to export an Aspose.Cells Workbook to PDF format while handling errors and ensuring the output path is ready.
// Common Searches: C# Aspose.Cells example for converting Excel file to PDF with error handling | How to save a workbook as PDF using Aspose.Cells in a .NET console application | Aspose.Cells C# code to check input file and create output directory before PDF export | Convert .xlsx to .pdf programmatically with Aspose.Cells in Visual Studio
// Tags: Aspose.Cells workbook to PDF conversion C# | C# console Aspose.Cells PDF export | validate input file before Aspose.Cells save | create output directory Aspose.Cells PDF

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookConversion
{
    // The console application verifies the presence of the source .xlsx file, ensures the target folder exists, loads the workbook with Aspose.Cells, and saves it directly as a PDF, wrapping the process in robust exception handling.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for input workbook and final PDF
            string workbookPath = @"C:\Input\Sample.xlsx";
            string pdfPath = @"C:\Output\Sample.pdf";

            try
            {
                // Verify input workbook exists
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Input workbook not found: {workbookPath}");
                    return;
                }

                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(pdfPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Load the workbook using Aspose.Cells
                Workbook workbook = new Workbook(workbookPath);

                // Directly save the workbook as PDF
                workbook.Save(pdfPath, SaveFormat.Pdf);

                Console.WriteLine("Conversion to PDF completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
