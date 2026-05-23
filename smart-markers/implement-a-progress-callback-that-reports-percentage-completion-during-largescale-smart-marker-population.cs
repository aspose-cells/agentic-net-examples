using System;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerProgressDemo
{
    // Callback implementation that reports percentage completion
    public class SmartMarkerProgressCallback : ISmartMarkerCallBack
    {
        private readonly int _totalMarkers;
        private int _processedMarkers;

        public SmartMarkerProgressCallback(int totalMarkers)
        {
            _totalMarkers = totalMarkers > 0 ? totalMarkers : 1; // avoid division by zero
            _processedMarkers = 0;
        }

        // This method is called for each smart marker that is processed
        public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
        {
            _processedMarkers++;

            // Calculate percentage (rounded to whole number)
            int percent = (int)((double)_processedMarkers / _totalMarkers * 100);

            Console.WriteLine($"Processed marker {_processedMarkers}/{_totalMarkers} " +
                              $"({percent}% ) - Sheet:{sheetIndex}, Row:{rowIndex}, Col:{colIndex}, " +
                              $"Table:{tableName}, Column:{columnName}");
        }
    }

    class Program
    {
        static void Main()
        {
            // Load the template workbook that contains smart markers
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = new Workbook("SmartMarkerTemplate.xlsx");

            // Prepare a sample data source (replace with your actual data)
            DataTable dt = new DataTable("Employees");
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Age", typeof(int));
            dt.Columns.Add("Department", typeof(string));

            // Populate many rows to simulate a large‑scale operation
            for (int i = 1; i <= 5000; i++)
            {
                dt.Rows.Add($"Employee {i}", 20 + (i % 30), $"Dept {(i % 5) + 1}");
            }

            // Bind the data source to the designer
            designer.SetDataSource("Employees", dt);

            // Determine total number of unique smart markers in the workbook
            int totalMarkers = designer.GetSmartMarkers().Length;

            // Assign the progress callback
            designer.CallBack = new SmartMarkerProgressCallback(totalMarkers);

            // Process all smart markers
            designer.Process();

            // Save the populated workbook
            designer.Workbook.Save("SmartMarkerResult.xlsx");
        }
    }
}