// Title: Export the first PivotTable’s data rows to a CSV file with proper quoting using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx workbook, refreshes the first PivotTable, reads its TableRange2 cells, and writes the rows to a .csv file with escaped commas and quotes using Aspose.Cells. | Show how to iterate over a PivotTable’s result range (TableRange2) and create a comma‑separated output that handles embedded line breaks in C# with Aspose.Cells. | Provide a method that checks for the existence of a PivotTable, refreshes it, and streams its data directly to a CSV file while applying CSV‑compliant quoting rules in .NET.
// Common Searches: aspocells c# export pivot table rows to csv with quoting | how to save pivot table result range as csv using Aspose.Cells | c# write TableRange2 of a pivot table to a comma separated file | export excel pivot data to csv handling commas and line breaks aspocells
// Tags: aspocells pivot table csv export | c# TableRange2 to csv | refresh pivot before csv export aspocells | csv quoting for pivot data aspocells | export visible pivot rows aspocells

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The example loads an Excel workbook, verifies a PivotTable exists on the first worksheet, refreshes it, obtains its TableRange2 (the data area), and writes each row to a CSV file. Cell values are escaped according to CSV rules, handling commas, quotes, and line breaks, resulting in a properly formatted comma‑separated file.
class ExportPivotToCsv
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "pivot_output.csv";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook containing the PivotTable
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one PivotTable
            if (sheet.PivotTables == null || sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("Error: No PivotTable found on the first worksheet.");
                return;
            }

            // Retrieve the first PivotTable on the worksheet
            PivotTable pivot = sheet.PivotTables[0];

            // Refresh the PivotTable data
            pivot.RefreshData();

            // Get the data range of the PivotTable (the area that contains the result rows)
            CellArea dataRange = pivot.TableRange2;

            // Open a StreamWriter to create the CSV file
            using (StreamWriter writer = new StreamWriter(outputPath))
            {
                // Loop through each row in the PivotTable data range
                for (int row = dataRange.StartRow; row <= dataRange.EndRow; row++)
                {
                    // Collect cell values for the current row
                    string[] values = new string[dataRange.EndColumn - dataRange.StartColumn + 1];

                    for (int col = dataRange.StartColumn; col <= dataRange.EndColumn; col++)
                    {
                        // Retrieve the cell value
                        object val = sheet.Cells[row, col].Value;
                        string cellText = val?.ToString() ?? string.Empty;

                        // Escape double quotes
                        if (cellText.Contains("\""))
                            cellText = cellText.Replace("\"", "\"\"");

                        // Enclose in quotes if the value contains commas, quotes, or line breaks
                        if (cellText.Contains(",") || cellText.Contains("\"") || cellText.Contains("\n"))
                            cellText = $"\"{cellText}\"";

                        values[col - dataRange.StartColumn] = cellText;
                    }

                    // Write the comma‑separated line to the CSV file
                    writer.WriteLine(string.Join(",", values));
                }
            }

            Console.WriteLine($"PivotTable data exported successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
