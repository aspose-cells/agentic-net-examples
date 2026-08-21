// Title: Aspose.Cells for .NET (C#): Export a pivot table that summarizes all formula results by worksheet
// Description: Load a workbook, extract every formula cell's evaluated value together with its sheet name, create a temporary data sheet, build a pivot table that groups and sums results per worksheet, refresh the pivot, and save the file using Aspose.Cells.
// Keywords: Aspose.Cells C# pivot table | export pivot table .NET | summarize formula results Aspose.Cells | group formulas by worksheet | collect formula values C# | Aspose.Cells workbook automation | pivot table from formula data
// Common Searches: Aspose.Cells create pivot table from all formulas | C# code to summarize formula results per sheet | export pivot table that groups formula outcomes by worksheet name | how to collect formula values and build a pivot in Aspose.Cells | Aspose.Cells .NET generate summary pivot of calculated cells
// Developer Intent: Produce a workbook that contains a pivot table aggregating the evaluated results of every formula cell, organized by the originating worksheet.
// Use Cases: Financial audit: quickly see total calculated values on each department sheet. | Consolidated reporting: display per‑sheet sums of KPI calculations for executive review. | Quality control: validate that each worksheet's formulas produce expected aggregate totals.
// AI Prompts: Generate C# code with Aspose.Cells that extracts all formula results from a workbook and creates a pivot table summing them by worksheet name. | Show how to modify the pivot to count formula cells instead of summing their values. | Explain how to export the workbook containing the pivot table to PDF while preserving the worksheet‑grouped summary.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsPivotExport
{
    // Load a workbook, extract every formula cell's evaluated value together with its sheet name, create a temporary data sheet, build a pivot table that groups and sums results per worksheet, refresh the pivot, and save the file using Aspose.Cells.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Load the source workbook; create an empty one if the file does not exist
                Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

                // Add a worksheet to collect raw formula data
                Worksheet dataSheet = workbook.Worksheets.Add("FormulaData");
                Cells dataCells = dataSheet.Cells;

                // Header row
                dataCells["A1"].PutValue("Worksheet");
                dataCells["B1"].PutValue("FormulaResult");

                int currentRow = 1; // zero‑based index; row 1 is the second row (after header)

                // Iterate through all worksheets in the workbook
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    // Skip the sheet used for data collection
                    if (ws.Name == dataSheet.Name) continue;

                    Cells cells = ws.Cells;
                    // Get the used range of the worksheet
                    AsposeRange usedRange = cells.MaxDisplayRange;
                    int startRow = usedRange.FirstRow;
                    int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                    int startCol = usedRange.FirstColumn;
                    int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                    for (int r = startRow; r <= endRow; r++)
                    {
                        for (int c = startCol; c <= endCol; c++)
                        {
                            Cell cell = cells[r, c];
                            // Identify formula cells
                            if (cell.IsFormula)
                            {
                                // Write worksheet name and evaluated result to the data sheet
                                dataCells[currentRow, 0].PutValue(ws.Name);
                                dataCells[currentRow, 1].PutValue(cell.Value);
                                currentRow++;
                            }
                        }
                    }
                }

                // Add a worksheet that will contain the pivot table
                Worksheet pivotSheet = workbook.Worksheets.Add("FormulaPivot");

                // Define the source data range for the pivot table (including header)
                int startDataRow = 2; // Excel rows are 1‑based; data starts after header
                int lastDataRow = currentRow + 1; // convert zero‑based index to Excel row number
                string sourceData = $"=FormulaData!${CellsHelper.ColumnIndexToName(0)}${startDataRow}:${CellsHelper.ColumnIndexToName(1)}${lastDataRow}";

                // Add the pivot table to the pivot sheet, starting at cell A3
                int pivotIndex = pivotSheet.PivotTables.Add(sourceData, "A3", "FormulasBySheet");
                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

                // Configure the pivot table: worksheet name as row field, formula result as data field (sum)
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Worksheet");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "FormulaResult");

                // Optional: display data in tabular form for clearer output
                pivotTable.ShowInTabularForm();

                // Refresh all pivot tables to ensure they reflect the latest data
                workbook.Worksheets.RefreshPivotTables();

                // Save the workbook with the new sheets and pivot table
                workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
