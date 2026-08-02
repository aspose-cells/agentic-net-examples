using System;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerProgressDemo
{
    // Callback implementation that reports percentage progress
    public class SmartMarkerProgressCallback : ISmartMarkerCallBack
    {
        private int _processedCount;
        private readonly int _totalCount;

        public SmartMarkerProgressCallback(int totalCount)
        {
            _totalCount = totalCount > 0 ? totalCount : 1; // avoid division by zero
            _processedCount = 0;
        }

        // This method is invoked for each smart marker occurrence during processing
        public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
        {
            _processedCount++;
            int percent = (int)((double)_processedCount / _totalCount * 100);
            Console.WriteLine($"Processed {_processedCount}/{_totalCount} ({percent}%) - Sheet:{sheetIndex} Row:{rowIndex} Col:{colIndex} Table:{tableName} Column:{columnName}");
        }
    }

    class Program
    {
        static void Main()
        {
            // Load the template workbook that contains smart markers
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = new Workbook("TemplateWithSmartMarkers.xlsx");

            // Prepare a sample data source (replace with your actual data)
            DataTable dt = new DataTable("Employees");
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Age", typeof(int));
            dt.Rows.Add("John Doe", 30);
            dt.Rows.Add("Jane Smith", 28);
            dt.Rows.Add("Bob Johnson", 45);
            designer.SetDataSource(dt);

            // Estimate total number of smart marker occurrences.
            // GetSmartMarkers returns unique markers; for demonstration we use its length as total.
            string[] uniqueMarkers = designer.GetSmartMarkers();
            int totalMarkers = uniqueMarkers.Length;

            // Assign the progress callback
            designer.CallBack = new SmartMarkerProgressCallback(totalMarkers);

            // Process all smart markers in the workbook
            designer.Process(true);

            // Save the populated workbook
            designer.Workbook.Save("PopulatedResult.xlsx");
        }
    }
}