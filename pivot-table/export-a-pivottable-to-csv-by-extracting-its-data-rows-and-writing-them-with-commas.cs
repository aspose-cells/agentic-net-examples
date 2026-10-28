// Title: Export an Excel PivotTable to CSV using Aspose.Cells for .NET (C# example)
// Common Searches: C# Aspose.Cells example to save pivot table data as CSV | How to write pivot table rows to a comma‑separated file with Aspose.Cells | Export used range of worksheet containing a PivotTable to CSV in .NET | Aspose.Cells refresh pivot table and export to CSV programmatically | Convert Excel pivot table to CSV with proper quoting using Aspose.Cells
// Tags: aspose.cells pivot table csv export | c# extract worksheet used range to csv | aspose.cells refresh pivot before export | csv field quoting with aspose.cells | programmatic pivot table to csv conversion

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The C# program loads an Excel workbook, checks for a PivotTable on the first worksheet, refreshes the PivotTable, determines the worksheet's used range, iterates through each cell to build CSV lines with proper quoting and escaping for commas and double quotes, and writes the result to an output CSV file while handling missing‑file and write errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.csv";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook containing the PivotTable
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one PivotTable on the sheet
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No PivotTable found.");
                return;
            }

            // Get the first PivotTable
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Refresh data (calculate is performed automatically on Save/Export)
            pivotTable.RefreshData();

            // Determine the used range of the worksheet (includes the PivotTable data)
            int startRow = worksheet.Cells.MinRow;
            int startColumn = worksheet.Cells.MinColumn;
            int rowCount = worksheet.Cells.MaxDataRow - startRow + 1;
            int columnCount = worksheet.Cells.MaxDataColumn - startColumn + 1;

            // Build CSV content from the used range
            StringBuilder csvBuilder = new StringBuilder();

            for (int i = 0; i < rowCount; i++)
            {
                for (int j = 0; j < columnCount; j++)
                {
                    Cell cell = worksheet.Cells[startRow + i, startColumn + j];
                    string text = cell.StringValue ?? string.Empty;

                    // Escape double quotes
                    if (text.Contains("\""))
                        text = text.Replace("\"", "\"\"");

                    // Enclose in quotes if the field contains a comma or quote
                    if (text.Contains(",") || text.Contains("\""))
                        text = $"\"{text}\"";

                    csvBuilder.Append(text);

                    if (j < columnCount - 1)
                        csvBuilder.Append(",");
                }
                csvBuilder.AppendLine();
            }

            // Write the CSV to a file with error handling
            try
            {
                File.WriteAllText(outputPath, csvBuilder.ToString());
                Console.WriteLine($"PivotTable exported to {outputPath}");
            }
            catch (Exception writeEx)
            {
                Console.WriteLine($"Failed to write CSV file: {writeEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
