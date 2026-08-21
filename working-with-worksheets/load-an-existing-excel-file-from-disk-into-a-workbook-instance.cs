// Title: Load an Existing Excel File into an Aspose.Cells Workbook (C#)
// Description: C# example that checks for a .xlsx file, creates a simple workbook with sample data if missing, then loads the file with Aspose.Cells, accesses the first worksheet, and prints its name and data row count.
// Keywords: Aspose.Cells load workbook C# | read Excel file Aspose.Cells | open existing .xlsx C# | first worksheet name Aspose | create placeholder workbook Aspose | FileNotFoundException handling Aspose.Cells
// Common Searches: how to open an existing Excel file with Aspose.Cells in C# | Aspose.Cells create workbook if file not found | get first sheet name and row count using Aspose.Cells | C# load .xlsx and handle missing file Aspose
// Developer Intent: Load a spreadsheet from disk into a Workbook object and retrieve basic information from its first worksheet.
// Use Cases: Display the name and row count of the first sheet in a user‑provided Excel file. | Automatically generate a default workbook with sample data when the expected file is absent. | Validate and safely open Excel files in .NET applications using Aspose.Cells.
// AI Prompts: Generate C# code that uses Aspose.Cells to open a .xlsx file, creates it with sample data if it does not exist, and prints the first worksheet name and number of data rows. | Show how to catch FileNotFoundException when loading an Excel workbook with Aspose.Cells in C#. | Demonstrate creating a new workbook, adding sample data, saving it, and then reloading it using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsLoadExample
{
    // C# example that checks for a .xlsx file, creates a simple workbook with sample data if missing, then loads the file with Aspose.Cells, accesses the first worksheet, and prints its name and data row count.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file
            string filePath = @"C:\Data\Sample.xlsx";

            try
            {
                // Ensure the file exists; create a simple workbook if it does not
                if (!File.Exists(filePath))
                {
                    // Create directory if needed
                    string dir = Path.GetDirectoryName(filePath);
                    if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                    {
                        Directory.CreateDirectory(dir);
                    }

                    // Create a new workbook with a default sheet and sample data
                    Workbook newWb = new Workbook();
                    Worksheet newSheet = newWb.Worksheets[0];
                    newSheet.Name = "Sheet1";
                    newSheet.Cells["A1"].PutValue("Sample Data");
                    newWb.Save(filePath);
                    Console.WriteLine($"Created sample workbook at: {filePath}");
                }

                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Console.WriteLine($"Loaded workbook: {filePath}");
                Console.WriteLine($"First worksheet name: {sheet.Name}");
                Console.WriteLine($"Number of rows with data: {sheet.Cells.MaxDataRow + 1}");
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.WriteLine($"File not found: {fnfEx.FileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
