// Title: Load an Excel workbook from a file path and obtain the first worksheet with Aspose.Cells for .NET (C#)
// AI Prompts: Load a workbook from a given file path using Aspose.Cells, create the file if it does not exist, and return the initial worksheet. | Write a value to cell A1 of the obtained worksheet and save the workbook back to the same location with Aspose.Cells in C#.
// Common Searches: C# Aspose.Cells load workbook from specific path and get first sheet | How to create a new Excel file when it doesn't exist using Aspose.Cells .NET | Aspose.Cells write to cell A1 after opening or creating workbook in C#
// Tags: initialize workbook from filesystem Aspose.Cells | retrieve initial worksheet Aspose.Cells | auto-generate workbook when absent Aspose.Cells | populate cell A1 Aspose.Cells | commit workbook changes Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example checks whether a .xlsx file exists, loads it with Aspose.Cells (or creates a new workbook if missing), accesses the first worksheet, writes "Hello Aspose.Cells!" to cell A1, and saves the workbook back to the original path.
    class Program
    {
        static void Main(string[] args)
        {
            // Specify the path to the Excel file
            string filePath = @"C:\Path\To\Your\File.xlsx";

            Workbook workbook = null;

            try
            {
                if (File.Exists(filePath))
                {
                    // Load the workbook from the existing file
                    workbook = new Workbook(filePath);
                }
                else
                {
                    // Create a new workbook if the file does not exist
                    workbook = new Workbook();
                    workbook.Save(filePath);
                    Console.WriteLine($"File not found. Created a new workbook at: {filePath}");
                }

                // Access the first worksheet (index 0)
                Worksheet firstWorksheet = workbook.Worksheets[0];

                // Example operation: write a value to cell A1
                firstWorksheet.Cells["A1"].PutValue("Hello Aspose.Cells!");

                // Save any changes made to the workbook
                workbook.Save(filePath);
                Console.WriteLine("Workbook processed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
