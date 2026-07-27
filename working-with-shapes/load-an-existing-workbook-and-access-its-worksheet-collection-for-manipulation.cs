// Title: Load an Excel workbook, access its worksheets, edit a cell, and save using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to open an existing .xlsx file with Aspose.Cells' Workbook constructor, retrieve the WorksheetCollection, write a string to cell A1 of the first sheet, handle missing files and runtime errors, and save the modified workbook to a new location on Windows platforms.
// Keywords: Aspose.Cells load workbook C# | access worksheet collection Aspose.Cells | write to cell A1 Aspose.Cells | save modified Excel file .NET | Workbook constructor Aspose.Cells | error handling Aspose.Cells C# | Windows Excel automation Aspose
// Common Searches: how to open an existing Excel file with Aspose.Cells C# | Aspose.Cells example to modify cell A1 after loading workbook | save a changed workbook to a different path using Aspose.Cells | C# code for loading workbook and accessing worksheets Aspose
// Developer Intent: Open a workbook, retrieve its worksheets, change a cell value, and write the updated file back to disk.
// Use Cases: Populate a title cell in a template workbook before generating a report. | Apply a common header across all sheets after loading a multi‑sheet workbook. | Read data from the first worksheet, update specific cells, and export the result to a new file.
// AI Prompts: Create C# code that uses Aspose.Cells to open an existing .xlsx file, write "Hello from Aspose.Cells" to cell A1 of the first worksheet, and save the workbook to a new file with comprehensive error handling. | Provide a robust Aspose.Cells for .NET example that loads a workbook, accesses the WorksheetCollection, updates a cell, and gracefully handles FileNotFoundException and other runtime exceptions.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsLoadAndManipulate
{
    // Demonstrates how to open an existing .xlsx file with Aspose.Cells' Workbook constructor, retrieve the WorksheetCollection, write a string to cell A1 of the first sheet, handle missing files and runtime errors, and save the modified workbook to a new location on Windows platforms.
    public class LoadAndAccessWorksheets
    {
        public static void Run()
        {
            // Path to the folder containing the Excel file.
            string dataDir = @"C:\Data\"; // Adjust this path as needed.

            // Build full file paths.
            string inputPath = Path.Combine(dataDir, "input.xlsx");
            string outputPath = Path.Combine(dataDir, "output.xlsx");

            // Verify that the input file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load an existing workbook from file using the Workbook(string) constructor.
                Workbook workbook = new Workbook(inputPath);

                // Access the worksheet collection.
                WorksheetCollection worksheets = workbook.Worksheets;

                // Example manipulation: get the first worksheet and write a value to cell A1.
                Worksheet firstSheet = worksheets[0];
                firstSheet.Cells["A1"].PutValue("Hello from Aspose.Cells");

                // Save the modified workbook to a new file.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime exceptions.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application.
    public class Program
    {
        public static void Main(string[] args)
        {
            LoadAndAccessWorksheets.Run();
        }
    }
}
