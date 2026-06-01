using System;
using Aspose.Cells;

namespace SmartMarkerTimingDemo
{
    // Callback implementation that logs start and end timestamps of smart marker processing
    public class SmartMarkerTimerCallback : ISmartMarkerCallBack
    {
        private readonly int _totalMarkers;          // Total number of smart markers to be processed
        private int _processedCount;                 // Counter for processed markers
        private DateTime _startTime;                 // Timestamp when processing starts
        private DateTime _endTime;                   // Timestamp when processing ends

        public SmartMarkerTimerCallback(int totalMarkers)
        {
            _totalMarkers = totalMarkers;
            _processedCount = 0;
        }

        // This method is invoked for each smart marker during processing
        public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
        {
            // Record start time on the first invocation
            if (_processedCount == 0)
            {
                _startTime = DateTime.Now;
                Console.WriteLine($"Smart marker processing started at {_startTime:O}");
            }

            _processedCount++;

            // Optional: log each individual marker processing details
            Console.WriteLine($"Processing marker {_processedCount}/{_totalMarkers} - Sheet:{sheetIndex}, Row:{rowIndex}, Column:{colIndex}, Table:{tableName}, Column:{columnName}");

            // When the last marker is processed, record end time and output total duration
            if (_processedCount == _totalMarkers)
            {
                _endTime = DateTime.Now;
                Console.WriteLine($"Smart marker processing ended at {_endTime:O}");
                Console.WriteLine($"Total execution time: {_endTime - _startTime}");
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Load the template workbook that contains smart markers
            Workbook templateWorkbook = new Workbook("SmartMarkerTemplate.xlsx");

            // Initialize WorkbookDesigner with the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = templateWorkbook
            };

            // Example data source (replace with actual data as needed)
            System.Data.DataTable dataTable = new System.Data.DataTable("Employees");
            dataTable.Columns.Add("Name", typeof(string));
            dataTable.Columns.Add("Age", typeof(int));
            dataTable.Rows.Add("John Doe", 30);
            dataTable.Rows.Add("Jane Smith", 28);
            designer.SetDataSource(dataTable);

            // Determine total number of smart markers to be processed
            string[] smartMarkers = designer.GetSmartMarkers();
            int totalMarkers = smartMarkers.Length;

            // Assign the custom callback that measures processing time
            designer.CallBack = new SmartMarkerTimerCallback(totalMarkers);

            // Process all smart markers
            designer.Process();

            // Save the resulting workbook
            designer.Workbook.Save("SmartMarkerOutput.xlsx");
        }
    }
}