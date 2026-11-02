// Title: Export cells with data validation from an Excel worksheet to a JSON array using Aspose.Cells for .NET
// Common Searches: how to retrieve only data‑validated cells from an .xlsx file with Aspose.Cells in C# | C# Aspose.Cells example to export validated cell addresses and values to JSON | serialize Excel cells that have validation rules to JSON using Aspose.Cells .NET | extract cells with validation from first worksheet and write JSON array in C# | Aspose.Cells get validation object for each cell and output as JSON
// Tags: Aspose.Cells export validated cells JSON | C# extract data validation cells Aspose.Cells | serialize Excel validation data to JSON | iterate worksheet used range Aspose.Cells | retrieve cell validation objects C#

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// Loads an Excel workbook, scans the first worksheet's used range for cells that have data validation, collects each cell's address and value, serializes the collection into a formatted JSON array, and writes the JSON to the console.
class ExportValidatedCells
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Prepare a list to hold validated cell information
            var validatedCells = new List<Dictionary<string, object>>();

            // Determine the used range of the worksheet
            Aspose.Cells.Range usedRange = worksheet.Cells.MaxDisplayRange;

            // Iterate through each cell in the used range
            for (int row = usedRange.FirstRow; row < usedRange.FirstRow + usedRange.RowCount; row++)
            {
                for (int col = usedRange.FirstColumn; col < usedRange.FirstColumn + usedRange.ColumnCount; col++)
                {
                    Cell cell = worksheet.Cells[row, col];

                    // Retrieve validation object for the cell, if any
                    Validation validation = cell.GetValidation();

                    // If the cell has validation applied, store its address and value
                    if (validation != null)
                    {
                        var cellInfo = new Dictionary<string, object>
                        {
                            ["Address"] = cell.Name,
                            ["Value"] = cell.Value
                        };
                        validatedCells.Add(cellInfo);
                    }
                }
            }

            // Serialize the list of validated cells to a JSON array
            string jsonOutput = JsonSerializer.Serialize(validatedCells, new JsonSerializerOptions { WriteIndented = true });

            // Output the JSON string
            Console.WriteLine(jsonOutput);
        }
        catch (Exception ex)
        {
            // Log unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
