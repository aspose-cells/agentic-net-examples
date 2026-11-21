// Title: Export Excel formula dependent cells to CSV using Aspose.Cells GetDependents in C#
// Common Searches: C# Aspose.Cells get dependent cells of a formula and save to CSV | How to list cells that depend on A1 with Aspose.Cells | Export cell dependency list from workbook to CSV using Aspose.Cells | Retrieve all dependent cells for a source cell in .NET Excel library | Aspose.Cells GetDependents example writing results to a file
// Tags: Aspose.Cells GetDependents to CSV | export dependent cell addresses Aspose.Cells | list formula dependents C# | Excel cell dependency extraction .NET | write cell dependency list to CSV C#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook, selects cell A1 on the first worksheet, uses GetDependents(true) to obtain every cell that references the formula, and writes each dependent's worksheet name and address to a CSV file with proper escaping, while handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "dependents.csv";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Select the first worksheet and the source cell (adjust as needed)
            Worksheet worksheet = workbook.Worksheets[0];
            Cell sourceCell = worksheet.Cells["A1"];

            // Get dependent cells for the source cell (true = get all dependents)
            Cell[] dependentCells = sourceCell.GetDependents(true);

            // Prepare CSV content
            List<string> csvLines = new List<string>
            {
                "SheetName,CellName" // header
            };

            // If there are dependents, enumerate them
            if (dependentCells != null)
            {
                foreach (Cell depCell in dependentCells)
                {
                    string sheetName = depCell.Worksheet.Name;
                    string cellName = depCell.Name; // e.g., "B2"
                    csvLines.Add($"{EscapeCsv(sheetName)},{EscapeCsv(cellName)}");
                }
            }

            // Export the list to a CSV file
            File.WriteAllLines(outputPath, csvLines);
            Console.WriteLine($"Dependents exported successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper to escape CSV fields when they contain commas, quotes, or newlines
    static string EscapeCsv(string field)
    {
        if (field.Contains(",") || field.Contains("\"") || field.Contains("\n"))
        {
            field = field.Replace("\"", "\"\"");
            return $"\"{field}\"";
        }
        return field;
    }
}
