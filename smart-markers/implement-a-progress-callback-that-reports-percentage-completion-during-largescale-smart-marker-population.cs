// Title: How to implement a smart‑marker progress callback that reports percentage completion in Aspose.Cells for .NET
// AI Prompts: Create a C# class that implements ISmartMarkerCallBack and writes the current marker index and completion percentage to the console during WorkbookDesigner.Process. | Show C# code to retrieve the total number of smart markers from WorkbookDesigner, assign the custom callback, and invoke the processing of all markers. | Provide a Windows Forms example where the ISmartMarkerCallBack updates a ProgressBar control instead of using Console.WriteLine.
// Common Searches: Aspose.Cells .NET how to monitor smart marker processing progress | C# get total smart marker count before calling WorkbookDesigner.Process | example of ISmartMarkerCallBack showing percentage completed | display smart marker population progress in a UI using Aspose.Cells | track large smart marker population performance with a callback in Aspose.Cells
// Tags: ISmartMarkerCallBack percentage progress | WorkbookDesigner smart marker count | smart marker processing callback .NET | Aspose.Cells progress reporting for smart markers | console logging of smart marker population

using System;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerProgressDemo
{
    // Callback implementation that reports percentage completion
    // Demonstrates implementing ISmartMarkerCallBack to log processed marker count and completion percentage, obtaining the total smart marker count, assigning the callback to WorkbookDesigner, processing the markers, and saving the populated workbook.
    public class SmartMarkerProgressCallback : ISmartMarkerCallBack
    {
        private readonly int _totalMarkers;
        private int _processedMarkers;

        public SmartMarkerProgressCallback(int totalMarkers)
        {
            _totalMarkers = totalMarkers > 0 ? totalMarkers : 1; // avoid division by zero
            _processedMarkers = 0;
        }

        // This method is called by Aspose.Cells for each smart marker processed
        public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
        {
            _processedMarkers++;
            int percent = (int)((double)_processedMarkers / _totalMarkers * 100);
            Console.WriteLine($"Processed {_processedMarkers}/{_totalMarkers} ({percent}%) - Sheet:{sheetIndex} Row:{rowIndex} Col:{colIndex} Table:{tableName} Column:{columnName}");
        }
    }

    class Program
    {
        static void Main()
        {
            // Load the template workbook that contains smart markers
            Workbook templateWorkbook = new Workbook("SmartMarkerTemplate.xlsx");

            // Initialize WorkbookDesigner and assign the workbook
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = templateWorkbook;

            // ----- Prepare a sample data source -----
            DataTable employeeTable = new DataTable("Employees");
            employeeTable.Columns.Add("Name", typeof(string));
            employeeTable.Columns.Add("Age", typeof(int));
            employeeTable.Rows.Add("John Doe", 30);
            employeeTable.Rows.Add("Jane Smith", 28);
            employeeTable.Rows.Add("Bob Johnson", 45);
            // Bind the data source to the designer
            designer.SetDataSource(employeeTable);
            // ----------------------------------------

            // Determine total number of smart markers before processing
            int totalMarkers = designer.GetSmartMarkers().Length;

            // Assign the progress callback
            designer.CallBack = new SmartMarkerProgressCallback(totalMarkers);

            // Process all smart markers (true = preserve unrecognized markers)
            designer.Process(true);

            // Save the populated workbook
            designer.Workbook.Save("SmartMarkerPopulated.xlsx");
        }
    }
}
