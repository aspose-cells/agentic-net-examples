// Title: C# – Stream Smart Markers with a Row‑Range LightCellsDataHandler to Reduce Memory in Aspose.Cells
// Description: Shows how to load an Excel template in streaming mode using a custom LightCellsDataHandler that streams only a defined row interval, creates a named range for those rows, binds a DataTable, processes the smart markers inside the range, and saves the workbook. A fallback creates a minimal template when the original file is missing.
// Keywords: Aspose.Cells | C# | smart markers | streaming mode | LightCellsDataHandler | row range | memory optimization | range‑based processing | WorkbookDesigner | large Excel files
// Common Searches: Aspose.Cells LightCellsDataHandler row range example | process smart markers in streaming mode C# | reduce memory usage when handling big Excel files with smart markers | create named range for smart markers Aspose.Cells | fallback template when Excel file not found Aspose.Cells
// Developer Intent: Load a workbook with LightCellsDataHandler and process smart markers only in a selected block of rows to keep memory usage low.
// Use Cases: Generate a paginated report where only rows 1000‑1999 contain data, streaming just that block to avoid loading the whole file. | Handle a massive template that stores smart markers in a specific section, processing only that section to produce a filtered output. | Automatically create a simple workbook with placeholder smart markers when the source template is missing, then run normal smart‑marker processing.
// AI Prompts: Extend the RangeLimitedHandler to also restrict processing to a column interval while staying in streaming mode. | Provide code that processes multiple named smart‑marker ranges in one workbook using WorkbookDesigner.Process with LightCellsDataHandler. | Explain the impact of returning true from IsGatherString() in the custom handler and how it interacts with row‑range filtering.

using System;
using System.Data;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace SmartMarkerStreamingDemo
{
    // Custom LightCellsDataHandler that processes only rows within a specified range.
    // Shows how to load an Excel template in streaming mode using a custom LightCellsDataHandler that streams only a defined row interval, creates a named range for those rows, binds a DataTable, processes the smart markers inside the range, and saves the workbook. A fallback creates a minimal template when the original file is missing.
    class RangeLimitedHandler : LightCellsDataHandler
    {
        private readonly int _startRow; // inclusive, zero‑based
        private readonly int _endRow;   // inclusive, zero‑based

        public RangeLimitedHandler(int startRow, int endRow)
        {
            _startRow = startRow;
            _endRow = endRow;
        }

        // Called for each worksheet. Return true to continue processing this sheet.
        public bool StartSheet(Worksheet sheet) => true;

        // Called for each row. Return true only for rows inside the desired range.
        public bool StartRow(int rowIndex) => rowIndex >= _startRow && rowIndex <= _endRow;

        // Called after a row is started. Return true to continue processing its cells.
        public bool ProcessRow(Row row) => true;

        // Called for each cell in a row that is being processed.
        public bool StartCell(int columnIndex) => true;

        // Called for each cell that is being processed.
        public bool ProcessCell(Cell cell) => true;

        // Determines whether string values should be gathered into a global pool.
        public bool IsGatherString() => false;
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the template workbook that contains smart markers.
                const string templatePath = "Template.xlsx";

                // Verify that the template file exists; if not, create a minimal workbook.
                Workbook workbook;
                if (File.Exists(templatePath))
                {
                    // Define the row range (e.g., rows 1000‑1999) that we want to load and process.
                    const int startRow = 999; // zero‑based index for row 1000
                    const int endRow = 1998;  // zero‑based index for row 1999
                    const int rowCount = endRow - startRow + 1;

                    // Set up LoadOptions with the custom LightCellsDataHandler to stream only the required rows.
                    var loadOptions = new LoadOptions
                    {
                        LightCellsDataHandler = new RangeLimitedHandler(startRow, endRow)
                    };

                    // Load the workbook using the streaming options.
                    workbook = new Workbook(templatePath, loadOptions);

                    // Create a range that covers the same rows and columns where smart markers reside.
                    // Example assumes smart markers are in columns A and B.
                    Worksheet sheet = workbook.Worksheets[0];
                    AsposeRange smartMarkerRange = sheet.Cells.CreateRange(startRow, 0, rowCount, 2);
                    smartMarkerRange.Name = "_CellsSmartMarkers"; // Required name for range‑based processing.

                    // Prepare a simple data source (DataTable) that matches the smart marker fields.
                    DataTable data = new DataTable("MyData");
                    data.Columns.Add("Name", typeof(string));
                    data.Columns.Add("Value", typeof(double));

                    // Populate the data table with sample rows.
                    for (int i = 0; i < rowCount; i++)
                    {
                        data.Rows.Add($"Item {i + 1}", (i + 1) * 10.5);
                    }

                    // Set up the WorkbookDesigner with the loaded workbook.
                    var designer = new WorkbookDesigner
                    {
                        Workbook = workbook
                        // LineByLine is obsolete; range‑based processing is used instead.
                    };

                    // Bind the data source to the designer.
                    designer.SetDataSource(data);

                    // Process only the defined smart marker range.
                    designer.Process(smartMarkerRange, true);
                }
                else
                {
                    // If the template is missing, create a new workbook with placeholder smart markers.
                    workbook = new Workbook();
                    Worksheet sheet = workbook.Worksheets[0];
                    sheet.Name = "Data";

                    // Insert simple smart markers in the first two columns.
                    sheet.Cells["A1"].PutValue("&=Name");
                    sheet.Cells["B1"].PutValue("&=Value");

                    // Prepare a minimal data source.
                    DataTable data = new DataTable("MyData");
                    data.Columns.Add("Name", typeof(string));
                    data.Columns.Add("Value", typeof(double));
                    data.Rows.Add("Sample", 123.45);

                    var designer = new WorkbookDesigner
                    {
                        Workbook = workbook
                    };
                    designer.SetDataSource(data);
                    designer.Process(true);
                }

                // Save the processed workbook.
                const string outputPath = "Output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
