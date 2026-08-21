// Title: Log Smart Marker Merges with ISmartMarkerCallBack and a Custom File Logger in Aspose.Cells for .NET
// Description: Shows how to attach a SimpleLogger to the WorkbookDesigner notify callback, capture every successful smart‑marker merge, write a timestamped entry to a text file, and save the processed workbook.
// Keywords: Aspose.Cells | smart markers | ISmartMarkerCallBack | notify parameter | custom file logger | C# .NET | WorkbookDesigner logging | record merge events | debug smart marker processing | audit workbook generation
// Common Searches: Aspose.Cells log smart marker merges | ISmartMarkerCallBack example C# | how to use notify parameter with Aspose.Cells | custom logger for smart markers .NET | track smart marker processing in Excel reports
// Developer Intent: Add a callback that writes a log entry for each smart‑marker merge during workbook processing.
// Use Cases: Create an audit trail of data rows merged into a report for compliance purposes. | Debug complex templates by reviewing which markers were processed and where they were placed. | Integrate merge‑event logging with existing monitoring or alerting systems.
// AI Prompts: Generate C# code that logs smart‑marker merges to a database instead of a text file. | Provide a thread‑safe logger implementation for high‑volume smart marker processing. | Show how to filter the notify callback to log only failed merges while still handling successful ones.

using System;
using System.Data;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Markup;

namespace SmartMarkerLoggingDemo
{
    // Simple logger that writes messages to a text file
    // Shows how to attach a SimpleLogger to the WorkbookDesigner notify callback, capture every successful smart‑marker merge, write a timestamped entry to a text file, and save the processed workbook.
    public class SimpleLogger
    {
        private readonly string _logFilePath;

        public SimpleLogger(string logFilePath)
        {
            _logFilePath = logFilePath;
            // Ensure the log file is empty at start
            File.WriteAllText(_logFilePath, string.Empty);
        }

        public void Log(string message)
        {
            string entry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}{Environment.NewLine}";
            File.AppendAllText(_logFilePath, entry);
        }
    }

    // Callback implementation that logs each smart‑marker processing event
    public class SmartMarkerLogger : ISmartMarkerCallBack
    {
        private readonly SimpleLogger _logger;

        public SmartMarkerLogger(SimpleLogger logger)
        {
            _logger = logger;
        }

        // This method is invoked by Aspose.Cells for every smart marker that is merged
        public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
        {
            _logger.Log($"Merged smart marker - Sheet:{sheetIndex}, Row:{rowIndex}, Column:{colIndex}, Table:{tableName}, Column:{columnName}");
        }
    }

    class Program
    {
        static void Main()
        {
            // Load a workbook that contains smart markers (template.xlsx should exist)
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = new Workbook("template.xlsx");

            // Prepare a simple data source
            DataTable data = new DataTable("Employees");
            data.Columns.Add("Name", typeof(string));
            data.Columns.Add("Age", typeof(int));
            data.Rows.Add("John Doe", 30);
            data.Rows.Add("Jane Smith", 28);

            // Assign the data source to the designer
            designer.SetDataSource(data);

            // Initialize the logger and assign the callback
            SimpleLogger logger = new SimpleLogger("SmartMarkerMergeLog.txt");
            designer.CallBack = new SmartMarkerLogger(logger);

            // Process all smart markers in the workbook
            designer.Process();

            // Save the resulting workbook
            designer.Workbook.Save("Result.xlsx");
        }
    }
}
