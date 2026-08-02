using System;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerTimingDemo
{
    // Callback implementation that records the first and last processing timestamps
    public class TimingSmartMarkerCallback : ISmartMarkerCallBack
    {
        private DateTime? _startTime;
        private DateTime? _endTime;

        // This method is called for each smart marker during processing
        public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
        {
            // Record start time on the first invocation
            if (_startTime == null)
                _startTime = DateTime.Now;

            // Update end time on every invocation (will hold the time of the last call)
            _endTime = DateTime.Now;

            // Optional: log each smart marker processing event
            Console.WriteLine($"SmartMarker processed - Sheet:{sheetIndex}, Row:{rowIndex}, Col:{colIndex}, Table:{tableName}, Column:{columnName}");
        }

        // Expose the measured duration
        public TimeSpan? GetProcessingDuration()
        {
            if (_startTime.HasValue && _endTime.HasValue)
                return _endTime - _startTime;
            return null;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Initialize the workbook designer with a template that contains smart markers
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = new Workbook("SmartMarkerTemplate.xlsx"); // replace with your template path

            // Set the custom callback to capture timing information
            TimingSmartMarkerCallback callback = new TimingSmartMarkerCallback();
            designer.CallBack = callback;

            // Prepare a simple data source (replace with your actual data)
            DataTable dt = new DataTable("Employees");
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Age", typeof(int));
            dt.Rows.Add("John Doe", 30);
            dt.Rows.Add("Jane Smith", 28);
            designer.SetDataSource(dt);

            // Process all smart markers
            designer.Process();

            // After processing, retrieve and display the total execution time
            TimeSpan? duration = callback.GetProcessingDuration();
            if (duration.HasValue)
                Console.WriteLine($"Total smart marker processing time: {duration.Value.TotalMilliseconds} ms");
            else
                Console.WriteLine("Processing duration could not be determined.");

            // Save the resulting workbook
            designer.Workbook.Save("SmartMarkerOutput.xlsx");
        }
    }
}