// Title: Convert all Excel files in a directory to PDF without charts using Aspose.Cells for .NET (C# command‑line tool)
// Common Searches: c# command line program to convert every .xlsx in a folder to PDF with Aspose.Cells | how to remove charts from Excel worksheets before exporting to PDF using Aspose.Cells | batch processing Excel files to PDF without charts in .NET | Aspose.Cells save workbook as PDF after clearing charts in C# | convert multiple Excel workbooks to PDF from a directory using Aspose.Cells CLI
// Tags: Aspose.Cells batch conversion to PDF | clear worksheet charts Aspose.Cells | C# CLI folder Excel processing | export workbook as PDF without charts | handle .xls .xlsx .xlsm programmatically

using System;
using System.IO;
using Aspose.Cells;

namespace ExcelToPdfConverter
{
    // A C# command‑line utility that takes a folder path, loads each .xls, .xlsx, or .xlsm workbook with Aspose.Cells, removes all charts from every worksheet, and saves the workbook as a PDF with the same name in the same directory.
    class Program
    {
        static void Main(string[] args)
        {
            // Verify that a folder path was provided
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: ExcelToPdfConverter <folderPath>");
                return;
            }

            string folderPath = args[0];

            // Check if the folder exists
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Error: The folder \"{folderPath}\" does not exist.");
                return;
            }

            // Get all Excel files in the folder (supports .xls, .xlsx, .xlsm)
            string[] excelFiles = Directory.GetFiles(folderPath, "*.xls*");

            if (excelFiles.Length == 0)
            {
                Console.WriteLine("No Excel files found in the specified folder.");
                return;
            }

            foreach (string excelFilePath in excelFiles)
            {
                try
                {
                    // Load the workbook
                    Workbook workbook = new Workbook(excelFilePath);

                    // Remove all charts from each worksheet
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        // Clear the Charts collection
                        sheet.Charts.Clear();
                    }

                    // Determine output PDF file path (same name, .pdf extension)
                    string pdfFilePath = Path.ChangeExtension(excelFilePath, ".pdf");

                    // Save the workbook as PDF
                    workbook.Save(pdfFilePath, SaveFormat.Pdf);

                    Console.WriteLine($"Converted \"{Path.GetFileName(excelFilePath)}\" to PDF successfully.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to convert \"{Path.GetFileName(excelFilePath)}\": {ex.Message}");
                }
            }
        }
    }
}
