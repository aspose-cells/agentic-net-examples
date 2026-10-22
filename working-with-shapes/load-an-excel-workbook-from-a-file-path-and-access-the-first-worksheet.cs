// Title: Load an Excel workbook from a file path and retrieve the first worksheet with Aspose.Cells for .NET (auto‑create if missing)
// AI Prompts: Open a workbook at a given path using Aspose.Cells, creating a new .xlsx file when the path does not exist, then obtain the worksheet at index 0. | Insert the text "Hello Aspose!" into cell A1 of the first worksheet and save the workbook back to the same location.
// Common Searches: asp.net how to open an existing .xlsx file with Aspose.Cells and get the first sheet | c# Aspose.Cells create workbook when file not found and write to cell A1 | load excel workbook from absolute path using Aspose.Cells and access worksheet index 0 | check file existence before loading Excel with Aspose.Cells in C#
// Tags: load workbook from file path Aspose.Cells | auto‑create workbook if missing Aspose.Cells | retrieve first worksheet index 0 Aspose.Cells | write string to cell A1 Aspose.Cells | save workbook after changes Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example checks whether a specified .xlsx file exists, loads it with Aspose.Cells or creates a new workbook if absent, accesses the first worksheet, writes "Hello Aspose!" to cell A1, and saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file (modify as needed)
            string filePath = @"C:\Path\To\Your\File.xlsx";

            Workbook workbook = null;

            try
            {
                if (File.Exists(filePath))
                {
                    // Load the existing workbook
                    workbook = new Workbook(filePath);
                }
                else
                {
                    // Create a new workbook if the file does not exist
                    workbook = new Workbook();
                    workbook.Worksheets[0].Name = "Sheet1";

                    // Ensure the target directory exists
                    string directory = Path.GetDirectoryName(filePath);
                    if (!Directory.Exists(directory))
                    {
                        Directory.CreateDirectory(directory);
                    }

                    // Save the newly created workbook
                    workbook.Save(filePath);
                    Console.WriteLine($"File not found. Created new workbook at '{filePath}'.");
                }

                // Access the first worksheet (index 0)
                Worksheet firstWorksheet = workbook.Worksheets[0];

                // Example operation: write a value to cell A1
                firstWorksheet.Cells["A1"].PutValue("Hello Aspose!");

                // Save any changes made to the workbook
                workbook.Save(filePath);
                Console.WriteLine("Workbook processed successfully.");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
