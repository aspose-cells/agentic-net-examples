// Title: Stream smart markers with Aspose.Cells using a row‑limited LightCellsDataHandler to reduce memory usage in .NET
// AI Prompts: Create a LightCellsDataHandler that streams only rows 1000‑1999 while processing smart markers in a workbook. | Modify the example to process only worksheets whose name starts with "Data" before applying the row‑range handler. | Update the handler to enable string pooling by returning true from IsGatherString and explain its impact on memory consumption.
// Common Searches: how to use LightCellsDataHandler to load only specific rows for smart marker processing in Aspose.Cells .NET | stream large Excel files with smart markers without loading the whole workbook in memory | Aspose.Cells memory optimization example for processing smart markers in a row range | custom LightCellsDataHandler filtering worksheets by name while using smart markers
// Tags: Aspose.Cells LightCellsDataHandler row range streaming | smart markers memory efficient processing .NET | load Excel workbook streaming mode Aspose.Cells | custom worksheet filter LightCellsDataHandler | string pooling LightCellsDataHandler Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace SmartMarkerStreamingDemo
{
    // Custom handler that streams only a specific row range of each worksheet.
    // This reduces memory consumption because cells outside the range are not kept in memory.
    // Demonstrates loading an Excel template that contains smart markers in streaming mode with a custom LightCellsDataHandler that processes only rows 1000‑1999, binds a list of Employee objects, processes the defined smart marker range, and saves the result, thereby minimizing memory usage.
    public class RangeLimitedLightCellsDataHandler : LightCellsDataHandler
    {
        private readonly int _startRow; // inclusive
        private readonly int _endRow;   // inclusive
        private bool _processCurrentSheet;

        public RangeLimitedLightCellsDataHandler(int startRow, int endRow)
        {
            _startRow = startRow;
            _endRow = endRow;
        }

        // Called for each worksheet. Return true to process this sheet.
        public bool StartSheet(Worksheet sheet)
        {
            // For demo we process all sheets; you can filter by name or index here.
            _processCurrentSheet = true;
            return true;
        }

        // Called for each row index. Return true only for rows within the desired range.
        public bool StartRow(int rowIndex)
        {
            if (!_processCurrentSheet) return false;
            return rowIndex >= _startRow && rowIndex <= _endRow;
        }

        // Called after a row is started. No special processing needed.
        public bool ProcessRow(Row row)
        {
            // Row data can be inspected here if required.
            return true; // Continue processing cells of this row.
        }

        // Called for each cell column index in the current row.
        public bool StartCell(int columnIndex)
        {
            // Process all cells of the selected rows.
            return true;
        }

        // Called for each cell that passed the above checks.
        public bool ProcessCell(Cell cell)
        {
            // Example: just read the value to keep the streaming alive.
            var value = cell.Value;
            // No further action needed for streaming; the cell is now in memory.
            return true;
        }

        // Determines whether string values should be gathered into a global string pool.
        public bool IsGatherString()
        {
            return false;
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the template workbook that contains smart markers.
                const string templatePath = "template.xlsx";

                // Ensure the template exists; if not, create a minimal one with smart markers.
                if (!File.Exists(templatePath))
                {
                    var tempWb = new Workbook();
                    var ws = tempWb.Worksheets[0];
                    // Insert simple smart markers.
                    ws.Cells["A2"].PutValue("&=Employees.Name");
                    ws.Cells["B2"].PutValue("&=Employees.Age");
                    ws.Cells["C2"].PutValue("&=Employees.Department");
                    // Define the named range expected by the designer.
                    var tempRange = ws.Cells.CreateRange("A2:C2");
                    tempRange.Name = "_CellsSmartMarkers";
                    tempWb.Save(templatePath);
                }

                // Define the row range we want to load (e.g., rows 1000‑1999).
                const int startRow = 1000;
                const int endRow = 1999;

                // Configure load options to use the custom LightCellsDataHandler.
                var loadOptions = new LoadOptions
                {
                    LightCellsDataHandler = new RangeLimitedLightCellsDataHandler(startRow, endRow)
                };

                // Load the workbook in streaming mode. Only the specified rows are kept in memory.
                var workbook = new Workbook(templatePath, loadOptions);

                // Prepare a simple data source for the smart markers.
                var employees = new List<Employee>
                {
                    new Employee { Name = "Alice", Age = 30, Department = "HR" },
                    new Employee { Name = "Bob",   Age = 45, Department = "IT" }
                };

                // Set up the WorkbookDesigner to process smart markers.
                var designer = new WorkbookDesigner
                {
                    Workbook = workbook,
                    // When LineByLine is false the designer expects a named range "_CellsSmartMarkers".
                    LineByLine = false
                };

                // Bind the data source to a name used in the template (e.g., "Employees").
                designer.SetDataSource("Employees", employees);

                // Define the range that contains the smart markers.
                // The template must have a range that encloses the markers and is named "_CellsSmartMarkers".
                Aspose.Cells.Range smartMarkerRange = workbook.Worksheets[0].Cells.CreateRange("A2:C2");
                smartMarkerRange.Name = "_CellsSmartMarkers";

                // Process only the defined range. The 'true' flag preserves unrecognized markers.
                designer.Process(smartMarkerRange, true);

                // Save the processed workbook.
                workbook.Save("output.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Simple POCO used as a data source for demonstration.
    public class Employee
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
        public string Department { get; set; } = string.Empty;
    }
}
