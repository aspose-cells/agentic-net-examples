// Title: How to expand a named range and refresh all pivot tables in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# expand named range and refresh pivot tables | Update pivot table data source after adding rows to a named range in .NET | Programmatically increase Excel named range rows and recalculate pivots with Aspose | Refresh all pivot tables after modifying a named range using Aspose.Cells | C# change RefersTo of a named range and refresh associated pivot tables
// Tags: expand named range Aspose.Cells | pivot table refresh data Aspose.Cells | update RefersTo named range .NET | recalculate pivot tables programmatically | save workbook after pivot refresh Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // PivotTable resides in this namespace

namespace AsposeCellsExample
{
    // The example loads an existing workbook, retrieves the 'DataRange' named range, extends its end row, updates the RefersTo reference, iterates through every worksheet to refresh and recalculate each pivot table, and saves the modified workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "InputWorkbook.xlsx";
                const string outputPath = "OutputWorkbook.xlsx";

                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Retrieve the named range that serves as the data source for the pivot table
                Name dataRange = workbook.Worksheets.Names["DataRange"];
                if (dataRange == null)
                {
                    Console.WriteLine("Named range 'DataRange' not found.");
                    return;
                }

                // Extract sheet name and cell area from the RefersTo string (e.g., Sheet1!A1:C10)
                string refersTo = dataRange.RefersTo;
                string sheetName;
                string rangePart;

                if (refersTo.Contains("!"))
                {
                    var parts = refersTo.Split('!');
                    sheetName = parts[0];
                    rangePart = parts[1];
                }
                else
                {
                    // If no sheet name is present, default to the first worksheet
                    sheetName = workbook.Worksheets[0].Name;
                    rangePart = refersTo;
                }

                // Parse the cell area (expects format A1:C10 or a single cell like A1)
                CellArea area;
                string[] cells = rangePart.Split(':');
                if (cells.Length == 2)
                {
                    area = CellArea.CreateCellArea(cells[0], cells[1]);
                }
                else
                {
                    // Single cell range
                    area = CellArea.CreateCellArea(cells[0], cells[0]);
                }

                // Expand the range by adding 10 rows (adjust as needed)
                const int rowsToAdd = 10;
                area.EndRow += rowsToAdd;

                // Build the new RefersTo string with the original sheet name
                string startCell = CellsHelper.CellIndexToName(area.StartRow, area.StartColumn);
                string endCell = CellsHelper.CellIndexToName(area.EndRow, area.EndColumn);
                dataRange.RefersTo = $"{sheetName}!{startCell}:{endCell}";

                // Refresh all pivot tables that depend on the expanded named range
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (PivotTable pivotTable in sheet.PivotTables)
                    {
                        try
                        {
                            pivotTable.RefreshData();      // Refresh the data source
                            pivotTable.CalculateData();    // Recalculate the pivot table
                        }
                        catch (Exception ptEx)
                        {
                            Console.WriteLine($"Error refreshing pivot table '{pivotTable.Name}': {ptEx.Message}");
                        }
                    }
                }

                // Save the updated workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
