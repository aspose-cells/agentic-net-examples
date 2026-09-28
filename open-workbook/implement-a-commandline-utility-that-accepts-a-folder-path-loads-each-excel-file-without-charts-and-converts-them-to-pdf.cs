// Title: Create a C# console utility to batch‑convert Excel workbooks (.xls, .xlsx, .xlsm) to PDF while stripping all charts with Aspose.Cells
// AI Prompts: Write a C# console program that takes a folder path argument, loads every .xls* file with Aspose.Cells, clears the Charts collection of each worksheet, and saves the workbook as a PDF. | Generate code to enumerate Excel files in a directory, handle missing folder or no files, and export each workbook to PDF using Workbook.Save with SaveFormat.Pdf after removing charts. | Add comprehensive try‑catch blocks and user‑friendly messages to a batch Excel‑to‑PDF command‑line tool that removes charts before conversion.
// Common Searches: c# aspose.cells batch convert excel files to pdf without charts | how to remove all charts from worksheets before exporting to pdf in .net | command line tool for converting .xls, .xlsx, .xlsm to pdf using Aspose.Cells | process multiple Excel workbooks in a folder and save as pdf in C#
// Tags: batch convert Excel to PDF Aspose.Cells | remove worksheet charts Aspose.Cells | command line Excel PDF conversion C# | process .xls* files programmatically Aspose.Cells | save workbook as PDF without charts

using System;
using System.IO;
using Aspose.Cells;

namespace ExcelToPdfWithoutCharts
{
    // Console app that receives a folder path, finds all .xls, .xlsx, .xlsm files, loads each workbook with Aspose.Cells, clears every chart from all worksheets, and saves the result as a PDF with the same name.
    class Program
    {
        static void Main(string[] args)
        {
            // Verify that a folder path was provided
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: ExcelToPdfWithoutCharts <folderPath>");
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
                    Console.WriteLine($"Failed to process \"{Path.GetFileName(excelFile)}\": {ex.Message}");
                }
            }
        }
    }
}
