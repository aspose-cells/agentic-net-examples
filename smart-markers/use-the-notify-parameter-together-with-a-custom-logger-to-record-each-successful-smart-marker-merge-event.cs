using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Markup;

namespace AsposeCellsSmartMarkerLogging
{
    // Simple logger that records messages to the console.
    public class CustomLogger
    {
        public void Log(string message)
        {
            Console.WriteLine($"[SmartMarkerLog] {DateTime.Now:O} - {message}");
        }
    }

    // Callback implementation that is invoked for each smart marker processed.
    public class SmartMarkerLoggerCallback : ISmartMarkerCallBack
    {
        private readonly CustomLogger _logger;

        public SmartMarkerLoggerCallback(CustomLogger logger)
        {
            _logger = logger;
        }

        // This method is called by Aspose.Cells for every smart marker merge.
        public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
        {
            // Build a descriptive message and log it.
            string cellAddress = CellsHelper.CellIndexToName(rowIndex, colIndex);
            string message = $"Merged smart marker at Sheet[{sheetIndex}] Cell[{cellAddress}] " +
                             $"Table=\"{tableName}\" Column=\"{columnName}\".";
            _logger.Log(message);
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Initialize logger.
            CustomLogger logger = new CustomLogger();

            // Create a new workbook that will act as a template.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Insert smart markers into the template.
            // &=$Products.Name and &=$Products.Price will be replaced during processing.
            cells["A1"].PutValue("&=$Products.Name");
            cells["B1"].PutValue("&=$Products.Price");

            // Prepare a data source (DataTable) matching the smart markers.
            DataTable productTable = new DataTable("Products");
            productTable.Columns.Add("Name", typeof(string));
            productTable.Columns.Add("Price", typeof(double));
            productTable.Rows.Add("Apple", 1.20);
            productTable.Rows.Add("Banana", 0.80);
            productTable.Rows.Add("Cherry", 2.50);

            // Set up the WorkbookDesigner.
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook,
                // Assign the custom callback to capture merge events.
                CallBack = new SmartMarkerLoggerCallback(logger)
            };

            // Register the data source.
            designer.SetDataSource(productTable);

            // Process all smart markers in the workbook.
            designer.Process();

            // Save the resulting workbook.
            workbook.Save("SmartMarkerMergedOutput.xlsx");
        }
    }
}