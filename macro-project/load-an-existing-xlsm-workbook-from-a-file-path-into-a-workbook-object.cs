// Title: Loading a macro‑enabled XLSM workbook from a file path into an Aspose.Cells Workbook using C#
// Common Searches: c# asp.net load existing .xlsm file with Aspose.Cells | aspocells read macro enabled workbook from disk c# | how to check file existence before loading workbook Aspose.Cells C# | write value to cell A1 and save as Xlsm using Aspose.Cells C# | load and modify macro-enabled Excel file programmatically with Aspose.Cells
// Tags: load macro-enabled workbook Aspose.Cells C# | Workbook constructor file path Aspose.Cells | verify file existence Aspose.Cells C# | write cell value Aspose.Cells C# | save workbook as Xlsm Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example checks whether a specified .xlsm file exists, loads it into an Aspose.Cells Workbook via the file‑path constructor, writes a text string to cell A1 of the first worksheet, and then saves the workbook as a macro‑enabled XLSM file.
    class Program
    {
        static void Main()
        {
            // Path to the XLSM workbook you want to load
            string filePath = @"C:\Path\To\YourWorkbook.xlsm";

            try
            {
                Workbook workbook;

                if (File.Exists(filePath))
                {
                    // Load the existing workbook
                    workbook = new Workbook(filePath);
                    Console.WriteLine("Workbook loaded successfully.");
                }
                else
                {
                    // File not found – create a new workbook (will be saved as macro‑enabled later)
                    workbook = new Workbook();
                    workbook.Worksheets[0].Name = "Sheet1";
                    Console.WriteLine("File not found. Created a new workbook.");
                }

                // Example operation: write text to cell A1 of the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Hello Aspose.Cells!");

                // Save the workbook (optional)
                string outputPath = Path.Combine(Environment.CurrentDirectory, "Output.xlsm");
                try
                {
                    workbook.Save(outputPath, SaveFormat.Xlsm);
                    Console.WriteLine($"Workbook saved to: {outputPath}");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Error saving workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
