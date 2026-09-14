// Title: Load an Excel workbook from a file path and list all worksheet names with Aspose.Cells for .NET
// AI Prompts: Write C# code that uses Aspose.Cells to open an Excel file from a specified path, creates a new workbook if the file does not exist, and prints each worksheet's name to the console. | Show how to safely load a workbook, handle missing files by initializing a default workbook, and iterate over the Worksheets collection to display sheet names using Aspose.Cells in .NET.
// Common Searches: aspnet c# how to open an existing Excel file with Aspose.Cells and get worksheet names | aspose.cells load workbook from path and create new workbook when file missing | c# enumerate worksheets collection after loading workbook using Aspose.Cells | handle FileNotFoundException with Aspose.Cells by creating default workbook
// Tags: load workbook from file path Aspose.Cells | create default workbook when file missing Aspose.Cells | iterate worksheets collection C# Aspose.Cells | list worksheet names Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

namespace SampleApp
{
    // The sample loads an Excel workbook from a given file path using Aspose.Cells for .NET, creates a default workbook if the file is absent, then iterates through the Worksheets collection and writes each sheet name to the console, with basic exception handling.
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = @"C:\Data\SampleWorkbook.xlsx";

            try
            {
                Workbook workbook;

                // Load existing workbook or create a new one if the file is missing
                if (File.Exists(filePath))
                {
                    workbook = new Workbook(filePath);
                }
                else
                {
                    workbook = new Workbook(); // creates a default workbook with one sheet
                    workbook.Save(filePath);   // optional: persist for future runs
                }

                // Access worksheets collection
                WorksheetCollection worksheets = workbook.Worksheets;

                // Output each worksheet name
                foreach (Worksheet sheet in worksheets)
                {
                    Console.WriteLine(sheet.Name);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
