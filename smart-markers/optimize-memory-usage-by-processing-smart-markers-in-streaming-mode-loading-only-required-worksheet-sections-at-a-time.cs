using System;
using System.Data;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace SmartMarkerStreamingDemo
{
    // Custom LightCellsDataHandler that loads only rows within a specified range.
    public class SectionLoadingHandler : LightCellsDataHandler
    {
        private readonly int _startRow; // inclusive
        private readonly int _endRow;   // inclusive

        public SectionLoadingHandler(int startRow, int endRow)
        {
            _startRow = startRow;
            _endRow = endRow;
        }

        // Called for each worksheet; we process all worksheets.
        public bool StartSheet(Worksheet sheet) => true;

        // Called for each row index; return true only for rows we want to keep in memory.
        public bool StartRow(int rowIndex) => rowIndex >= _startRow && rowIndex <= _endRow;

        // ProcessRow can be used to inspect row data; we simply continue.
        public bool ProcessRow(Row row) => true;

        // Called for each cell column index; we process all cells in the selected rows.
        public bool StartCell(int columnIndex) => true;

        // ProcessCell can be used to inspect cell data; we simply continue.
        public bool ProcessCell(Cell cell) => true;
    }

    public class Program
    {
        // Sample method that creates a DataTable used as a smart‑marker data source.
        private static DataTable CreateSampleData()
        {
            var dt = new DataTable("Employees");
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Age", typeof(int));
            dt.Columns.Add("Department", typeof(string));

            dt.Rows.Add("John Doe", 30, "Sales");
            dt.Rows.Add("Jane Smith", 28, "Marketing");
            dt.Rows.Add("Bob Johnson", 35, "IT");

            return dt;
        }

        public static void Main()
        {
            try
            {
                // Define the rows that contain smart markers (e.g., rows 1000‑1099).
                const int smartMarkerStartRow = 1000;
                const int smartMarkerRowCount = 100; // number of rows with smart markers

                // Verify that the template file exists.
                const string templatePath = "SmartMarkerTemplate.xlsx";
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file not found: {templatePath}");
                    return;
                }

                // Set up LoadOptions with the custom LightCellsDataHandler to stream only the needed rows.
                var loadOptions = new LoadOptions
                {
                    LightCellsDataHandler = new SectionLoadingHandler(
                        smartMarkerStartRow,
                        smartMarkerStartRow + smartMarkerRowCount - 1)
                };

                // Load the template workbook in streaming mode.
                var workbook = new Workbook(templatePath, loadOptions);

                // Access the worksheet that contains the smart markers.
                var sheet = workbook.Worksheets[0];

                // Define a named range that encloses the smart markers.
                // CreateRange(startRow, startColumn, totalRows, totalColumns)
                AsposeRange smartMarkerRange = sheet.Cells.CreateRange(
                    smartMarkerStartRow, 0, smartMarkerRowCount, 5);
                smartMarkerRange.Name = "_CellsSmartMarkers";

                // Set up the WorkbookDesigner.
                var designer = new WorkbookDesigner
                {
                    Workbook = workbook,
                    // LineByLine is obsolete; kept for compatibility but not required.
                    LineByLine = false
                };

                // Assign the data source for the smart markers.
                designer.SetDataSource("Employees", CreateSampleData());

                // Process only the defined range; true = preserve unrecognized markers (not needed here).
                designer.Process(smartMarkerRange, true);

                // Save the processed workbook.
                const string outputPath = "ProcessedOutput.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Processing complete. Output saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}