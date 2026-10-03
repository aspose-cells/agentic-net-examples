// Title: Add a Timestamp Worksheet to a Merged Excel File Using Aspose.Cells for .NET
// AI Prompts: Create a new worksheet named "Timestamp" in the loaded workbook, insert DateTime.Now into cell A1, and apply the built‑in date‑time number format (22). | Save the updated workbook to a separate file (e.g., merged_with_timestamp.xlsx) while keeping the original merged workbook unchanged.
// Common Searches: Aspose.Cells C# add a new worksheet with current timestamp to an existing Excel file | how to format a cell as date‑time number using Aspose.Cells | save a copy of a workbook after adding extra sheets with Aspose.Cells .NET
// Tags: add timestamp worksheet Aspose.Cells | insert current datetime into Excel cell C# | date‑time number format Aspose.Cells | save workbook with additional sheet .NET

using Aspose.Cells;
using System;
using System.IO;

// Loads an existing merged.xlsx workbook, adds a "Timestamp" worksheet, writes the current date and time into cell A1 with number format 22 (mm-dd-yy hh:mm), and saves the result as merged_with_timestamp.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "merged.xlsx";
            const string outputPath = "merged_with_timestamp.xlsx";

            // Ensure the source workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook that resulted from the merge operation
            Workbook workbook = new Workbook(inputPath);

            // Add a new worksheet (returns its index)
            int sheetIndex = workbook.Worksheets.Add();
            Worksheet tsSheet = workbook.Worksheets[sheetIndex];
            tsSheet.Name = "Timestamp";

            // Write the current date and time into cell A1
            Cell cell = tsSheet.Cells["A1"];
            cell.PutValue(DateTime.Now);

            // Apply date‑time number format (22 = mm-dd-yy hh:mm)
            Style style = cell.GetStyle();
            style.Number = 22;
            cell.SetStyle(style);

            // Save the workbook with the timestamp worksheet
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
