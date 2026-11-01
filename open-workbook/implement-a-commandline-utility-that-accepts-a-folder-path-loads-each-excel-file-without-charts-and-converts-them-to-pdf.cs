// Title: Command‑line tool to batch convert Excel (.xls, .xlsx) workbooks to PDF while removing all charts using Aspose.Cells for .NET
// Common Searches: c# aspocells convert all excel files in a folder to pdf without charts | how to strip charts from worksheets before saving as pdf with Aspose.Cells | batch process .xls and .xlsx files to pdf via command line using Aspose.Cells | run excel to pdf conversion for multiple files in a directory with Aspose.Cells C# | aspocells command‑line utility to export workbooks to pdf after clearing charts
// Tags: batch excel to pdf conversion aspocells | remove worksheet charts aspocells | c# command line folder processing aspocells | save workbook as pdf without charts | process xls xlsx files programmatically aspocells

using System;
using System.IO;
using Aspose.Cells;

namespace ExcelToPdfConverter
{
    // A C# console application that accepts a folder path, enumerates all .xls and .xlsx files, loads each workbook with Aspose.Cells, clears every chart from every worksheet, and saves the result as a PDF with the same name. It reports success or errors for each file, enabling fast batch conversion of Excel workbooks to PDF without chart graphics.
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

            // Get all Excel files in the folder (including .xls and .xlsx)
            string[] excelFiles = Directory.GetFiles(folderPath, "*.xls*");

            if (excelFiles.Length == 0)
            {
                Console.WriteLine("No Excel files found in the specified folder.");
                return;
            }

            foreach (string excelFile in excelFiles)
            {
                try
                {
                    // Load the workbook
                    Workbook workbook = new Workbook(excelFile);

                    // Remove all charts from each worksheet
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        sheet.Charts.Clear();
                    }

                    // Determine the output PDF file path
                    string pdfFile = Path.ChangeExtension(excelFile, ".pdf");

                    // Save the workbook as PDF
                    workbook.Save(pdfFile, SaveFormat.Pdf);

                    Console.WriteLine($"Converted \"{Path.GetFileName(excelFile)}\" to PDF successfully.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to convert \"{Path.GetFileName(excelFile)}\": {ex.Message}");
                }
            }
        }
    }
}
