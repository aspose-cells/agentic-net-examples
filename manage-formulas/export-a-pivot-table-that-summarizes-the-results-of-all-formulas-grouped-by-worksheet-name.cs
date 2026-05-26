using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotExport
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Input and output file paths.
                string inputPath = "InputWithFormulas.xlsx";
                string outputPath = "OutputWithFormulaSummary.xlsx";

                // Verify that the input workbook exists.
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook that contains formulas.
                Workbook workbook = new Workbook(inputPath);

                // Create a worksheet to hold raw data for the pivot table.
                Worksheet dataSheet = workbook.Worksheets.Add("SummaryData");

                // Write header row.
                dataSheet.Cells[0, 0].PutValue("Worksheet");
                dataSheet.Cells[0, 1].PutValue("Cell");
                dataSheet.Cells[0, 2].PutValue("Value");

                int currentRow = 1; // Start after header.

                // Iterate through all worksheets and collect formula results.
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    // Skip the data sheet itself to avoid recursion.
                    if (ws.Name == dataSheet.Name)
                        continue;

                    Cells cells = ws.Cells;
                    // Enumerate all cells that contain a formula.
                    foreach (Cell cell in cells)
                    {
                        if (!string.IsNullOrEmpty(cell.Formula))
                        {
                            dataSheet.Cells[currentRow, 0].PutValue(ws.Name);
                            dataSheet.Cells[currentRow, 1].PutValue(cell.Name);
                            dataSheet.Cells[currentRow, 2].PutValue(cell.Value);
                            currentRow++;
                        }
                    }
                }

                // Add a new worksheet that will contain the pivot table.
                Worksheet pivotSheet = workbook.Worksheets.Add("FormulaSummary");

                // Define the source range for the pivot table (including header).
                string sourceRange = $"=SummaryData!A1:C{currentRow - 1}";

                // Add the pivot table to the pivot sheet.
                int pivotIndex = pivotSheet.PivotTables.Add(sourceRange, "A3", "FormulaPivot");
                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

                // Configure the pivot: group by worksheet name and summarize values.
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Worksheet");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Value"); // Default aggregation is Sum.

                // Optional: display the pivot in a tabular layout.
                pivotTable.ShowInTabularForm();

                // Refresh all pivot tables to ensure they reflect the latest data.
                workbook.Worksheets.RefreshPivotTables();

                // Save the workbook with the new summary and pivot table.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}