// Title: C# – Measure Smart Marker Processing Time with ISmartMarkerCallBack in Aspose.Cells
// Description: Demonstrates how to implement ISmartMarkerCallBack in Aspose.Cells for .NET to capture the first and last processing timestamps of smart markers, calculate total execution time, log each marker event, and save the workbook. Ideal for performance monitoring and benchmarking of large smart‑marker datasets.
// Keywords: Aspose.Cells | Smart Marker | ISmartMarkerCallBack | C# | .NET | performance timing | execution duration | callback logging | WorkbookDesigner | benchmark sample | GitHub
// Common Searches: Aspose.Cells smart marker timing example | ISmartMarkerCallBack record start end timestamps | measure smart marker processing duration C# | log each smart marker event with Aspose.Cells | benchmark Aspose.Cells smart markers
// Developer Intent: Obtain start and end timestamps of smart‑marker processing to compute total execution time and optionally log each marker call.
// Use Cases: Profile performance of complex smart‑marker templates. | Display processing duration in console or UI for diagnostics. | Integrate timing data into automated test suites or CI pipelines. | Compare execution times across different data volumes.
// AI Prompts: Create a C# snippet that uses ISmartMarkerCallBack to write processing timestamps to a CSV file. | Show how to extend SmartMarkerTimerCallback to collect per‑marker latency and output a summary report. | Provide code that formats the total smart‑marker duration as HH:mm:ss.fff and logs it with Serilog.

using System;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerTimingDemo
{
    // Callback implementation that records the first and last processing timestamps
    // Demonstrates how to implement ISmartMarkerCallBack in Aspose.Cells for .NET to capture the first and last processing timestamps of smart markers, calculate total execution time, log each marker event, and save the workbook. Ideal for performance monitoring and benchmarking of large smart‑marker datasets.
    public class SmartMarkerTimerCallback : ISmartMarkerCallBack
    {
        private DateTime _startTime;
        private DateTime _endTime;
        private bool _hasStarted = false;

        // Called for each smart marker during processing
        public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
        {
            // Record start time on the first invocation
            if (!_hasStarted)
            {
                _startTime = DateTime.Now;
                _hasStarted = true;
            }

            // Update end time on every invocation (will hold the time of the last call)
            _endTime = DateTime.Now;

            // Optional: log each smart marker processing event
            Console.WriteLine($"Processed marker at {DateTime.Now:O} - Sheet:{sheetIndex}, Row:{rowIndex}, Col:{colIndex}, Table:{tableName}, Column:{columnName}");
        }

        // Expose the captured timestamps
        public DateTime StartTime => _startTime;
        public DateTime EndTime => _endTime;
        public TimeSpan Duration => _endTime - _startTime;
    }

    public class Program
    {
        public static void Main()
        {
            // Load the template workbook that contains smart markers
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = new Workbook("SmartMarkerTemplate.xlsx");

            // Prepare a simple data source
            DataTable dt = new DataTable("Employees");
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Age", typeof(int));
            dt.Rows.Add("John Doe", 30);
            dt.Rows.Add("Jane Smith", 28);

            // Set the data source for the designer
            designer.SetDataSource(dt);

            // Attach the timing callback
            SmartMarkerTimerCallback callback = new SmartMarkerTimerCallback();
            designer.CallBack = callback;

            // Process all smart markers
            designer.Process();

            // After processing, output the total execution time captured by the callback
            Console.WriteLine($"Smart marker processing started at: {callback.StartTime:O}");
            Console.WriteLine($"Smart marker processing ended at:   {callback.EndTime:O}");
            Console.WriteLine($"Total processing duration: {callback.Duration.TotalMilliseconds} ms");

            // Save the resulting workbook
            designer.Workbook.Save("SmartMarkerOutput.xlsx");
        }
    }
}
