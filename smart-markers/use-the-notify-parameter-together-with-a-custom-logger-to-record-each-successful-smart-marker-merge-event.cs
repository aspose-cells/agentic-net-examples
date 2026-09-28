// Title: Use a custom ISmartMarkerCallBack with the notify parameter to log every smart marker merge in Aspose.Cells (C#)
// AI Prompts: Create a C# class that implements ISmartMarkerCallBack and writes sheet index, cell address, table and column names to a logger for each smart marker merge. | Show how to attach the callback to WorkbookDesigner, process a workbook with a DataTable source, and save the output while capturing merge events. | Demonstrate configuring a simple console logger and using CellsHelper to convert row/column indexes to cell names inside the callback.
// Common Searches: how to capture smart marker merge events with Aspose.Cells C# | example of ISmartMarkerCallBack implementation for logging in .NET | using notify parameter to track smart marker processing in Aspose.Cells | log smart marker merges to console while processing workbook designer | record smart marker merge details (sheet, cell, table) in Aspose.Cells
// Tags: custom ISmartMarkerCallBack merge logging | Aspose.Cells notify parameter callback usage | smart marker processing audit log .NET | WorkbookDesigner callback console logger | record smart marker merge details Excel

using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Markup;

namespace SmartMarkerLoggingDemo
{
    // Simple logger that records messages (here to console, could be extended to file, DB, etc.)
    // The example defines a CustomLogger that writes messages to the console, implements SmartMarkerLogger as an ISmartMarkerCallBack to log sheet index, cell address, table and column names for each smart marker merge, assigns this callback to WorkbookDesigner, processes a DataTable data source, and saves the resulting workbook while recording every merge event.
    public class CustomLogger
    {
        public void Log(string message)
        {
            // In a real scenario, replace this with proper logging infrastructure
            Console.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}");
        }
    }

    // Callback implementation that Aspose.Cells will invoke for each smart marker processing event
    public class SmartMarkerLogger : ISmartMarkerCallBack
    {
        private readonly CustomLogger _logger;

        public SmartMarkerLogger(CustomLogger logger)
        {
            _logger = logger;
        }

        // This method is called by the WorkbookDesigner during smart marker processing
        public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
        {
            // Build a descriptive message for the merge event
            string cellAddress = CellsHelper.CellIndexToName(rowIndex, colIndex);
            string message = $"Smart marker merged - Sheet:{sheetIndex}, Cell:{cellAddress}, Table:{tableName}, Column:{columnName}";
            _logger.Log(message);
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Initialize logger
            CustomLogger logger = new CustomLogger();

            // Load a workbook that contains smart markers (template.xlsx should exist in the execution folder)
            Workbook workbook = new Workbook("template.xlsx");

            // Create WorkbookDesigner and assign the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook,
                // Assign our custom callback so we get notified on each smart marker merge
                CallBack = new SmartMarkerLogger(logger)
            };

            // Prepare a simple data source (DataTable) matching the smart markers in the template
            System.Data.DataTable dataTable = new System.Data.DataTable("Employees");
            dataTable.Columns.Add("Name", typeof(string));
            dataTable.Columns.Add("Age", typeof(int));
            dataTable.Rows.Add("John Doe", 30);
            dataTable.Rows.Add("Jane Smith", 28);

            // Set the data source for the designer
            designer.SetDataSource(dataTable);

            // Process all smart markers; the callback will log each successful merge
            designer.Process();

            // Save the resulting workbook
            workbook.Save("output.xlsx");
        }
    }
}
