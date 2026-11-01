// Title: How to modify a cell value and auto‑fit its row using Aspose.Cells for .NET (C#)
// AI Prompts: Load an existing Excel workbook, change the value of a specific cell, invoke worksheet.AutoFitRow for that row, and save the workbook with Aspose.Cells in C#. | Write C# code that verifies the input file, updates a cell's content, automatically adjusts the row height to fit the new value, and writes the result to a new file using Aspose.Cells.
// Common Searches: Aspose.Cells C# auto fit row after updating cell value | C# change Excel cell and adjust row height programmatically with Aspose | How to use AutoFitRow on a specific row in Aspose.Cells .NET | Load workbook, modify cell, auto‑fit row, save using Aspose.Cells C# example | AutoFitRow method example for a single row in Aspose.Cells for .NET
// Tags: Aspose.Cells AutoFitRow C# | modify Excel cell value Aspose.Cells | auto‑fit row height after cell edit | load and save workbook Aspose.Cells .NET | row height adjustment based on content Aspose

using System;
using System.IO;
using Aspose.Cells;

// The sample loads input.xlsx, updates cell A1 with a new string, automatically adjusts the height of the row containing that cell using AutoFitRow, and saves the modified workbook as output.xlsx, including file existence checks and exception handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing workbook from file
            var workbook = new Workbook(inputPath);

            // Get the first worksheet (you can change the index or name as needed)
            var worksheet = workbook.Worksheets[0];

            // Modify the value of cell A1 (change the address as required)
            var cell = worksheet.Cells["A1"];
            cell.PutValue("New Value");

            // Auto‑fit the row that contains the modified cell
            int rowIndex = cell.Row; // zero‑based row index
            worksheet.AutoFitRow(rowIndex); // overload with a single argument

            // Save the workbook with the changes
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
