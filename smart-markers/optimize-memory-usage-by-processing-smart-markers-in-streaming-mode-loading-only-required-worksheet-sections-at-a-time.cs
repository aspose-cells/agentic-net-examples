// Title: C# – Process Smart Markers in Streaming Mode with LightCellsDataHandler (Row‑Range Loading)
// Description: Demonstrates how to use a custom LightCellsDataHandler to load only a specific row range from the first worksheet, define the smart‑marker range "_CellsSmartMarkers", bind a DataTable, and process smart markers with WorkbookDesigner in streaming mode, dramatically lowering memory consumption in Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | smart markers streaming | LightCellsDataHandler | partial row loading | memory optimization | WorkbookDesigner | named range _CellsSmartMarkers | large Excel processing
// Common Searches: Aspose.Cells LightCellsDataHandler limit rows | process smart markers without loading whole workbook | streaming smart markers C# example | reduce memory usage Aspose.Cells smart markers | named range _CellsSmartMarkers usage
// Developer Intent: Load a workbook in streaming mode, restrict memory to a defined row interval, and execute smart‑marker processing with WorkbookDesigner.
// Use Cases: Generate reports from massive Excel templates where smart markers occupy a known row block, avoiding full workbook loading. | Run server‑side batch jobs that populate smart markers on the first sheet while keeping RAM usage minimal. | Create a lightweight data‑binding routine that reads only required rows, processes smart markers, and writes the result back to disk.
// AI Prompts: Write C# code that uses LightCellsDataHandler to process smart markers located in rows 50‑100 of the first worksheet in streaming mode. | Explain how to configure WorkbookDesigner with LineByLine = false and the "_CellsSmartMarkers" named range for efficient smart‑marker processing. | Show how to extend RangeLimitedLightCellsDataHandler to handle multiple worksheets while still limiting memory usage.

using System;
using System.Data;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace SmartMarkerStreamingDemo
{
    // Custom handler to read only a specific row range from the worksheet.
    // This reduces memory consumption by skipping rows that are not needed.
    // Demonstrates how to use a custom LightCellsDataHandler to load only a specific row range from the first worksheet, define the smart‑marker range "_CellsSmartMarkers", bind a DataTable, and process smart markers with WorkbookDesigner in streaming mode, dramatically lowering memory consumption in Aspose.Cells for .NET.
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

        // Called for each worksheet. Process only the first sheet (index 0).
        public bool StartSheet(Worksheet sheet)
        {
            _processCurrentSheet = sheet.Index == 0;
            return _processCurrentSheet;
        }

        // Called for each row. Process only rows within the defined range.
        public bool StartRow(int rowIndex)
        {
            if (!_processCurrentSheet) return false;
            return rowIndex >= _startRow && rowIndex <= _endRow;
        }

        // Row processing can be used for additional logic; we simply continue.
        public bool ProcessRow(Row row) => true;

        // Called for each cell in a processed row. Process all cells.
        public bool StartCell(int columnIndex) => true;

        // Cell processing can be used for custom actions; we just continue.
        public bool ProcessCell(Cell cell) => true;

        // No need to gather strings into a global pool for this scenario.
        public bool IsGatherString() => false;
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the template workbook that contains smart markers.
                const string templatePath = "TemplateWithSmartMarkers.xlsx";

                // Ensure the template file exists; create a minimal one if missing.
                if (!File.Exists(templatePath))
                {
                    var tempWb = new Workbook();
                    var ws = tempWb.Worksheets[0];
                    // Insert example smart markers in the expected range (rows 10‑30).
                    ws.Cells["A10"].PutValue("&=Employees.Name");
                    ws.Cells["B10"].PutValue("&=Employees.Age");
                    ws.Cells["C10"].PutValue("&=Employees.Department");
                    tempWb.Save(templatePath);
                }

                // Define the row range that actually contains the smart markers.
                // For example, rows 10 to 30 (zero‑based indices 9 to 29).
                int smartMarkerStartRow = 9;
                int smartMarkerEndRow   = 29;

                // Set up load options with the custom LightCellsDataHandler.
                var loadOptions = new LoadOptions
                {
                    LightCellsDataHandler = new RangeLimitedLightCellsDataHandler(smartMarkerStartRow, smartMarkerEndRow)
                };

                // Load the workbook in streaming mode; only the specified rows are kept in memory.
                var workbook = new Workbook(templatePath, loadOptions);

                // Ensure the range that holds smart markers is named "_CellsSmartMarkers".
                // This is required when LineByLine is set to false.
                Worksheet sheet = workbook.Worksheets[0];
                AsposeRange smartMarkerRange = sheet.Cells.CreateRange(
                    smartMarkerStartRow,
                    0,
                    smartMarkerEndRow - smartMarkerStartRow + 1,
                    sheet.Cells.MaxDataColumn + 1);
                smartMarkerRange.Name = "_CellsSmartMarkers";

                // Prepare a simple data source (DataTable) that matches the smart markers.
                var dt = new DataTable("Employees");
                dt.Columns.Add("Name", typeof(string));
                dt.Columns.Add("Age", typeof(int));
                dt.Columns.Add("Department", typeof(string));
                dt.Rows.Add("John Doe", 30, "Sales");
                dt.Rows.Add("Jane Smith", 28, "Marketing");

                // Set up the WorkbookDesigner.
                var designer = new WorkbookDesigner
                {
                    Workbook = workbook,
                    LineByLine = false   // Process the named range instead of line‑by‑line.
                };

                // Bind the data source to the name used in the smart markers.
                designer.SetDataSource("Employees", dt);

                // Process the smart markers. Only the previously loaded rows are examined.
                designer.Process();

                // Save the populated workbook.
                const string outputPath = "ProcessedOutput.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Processing complete. Output saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
