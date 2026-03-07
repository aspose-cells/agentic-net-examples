using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace SmartMarkerNotificationDemo
{
    // Callback implementation to receive notifications while smart markers are being processed
    public class SmartMarkerNotifier : ISmartMarkerCallBack
    {
        // This method is called for each smart marker cell during processing
        public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
        {
            Console.WriteLine($"Processing smart marker - Sheet: {sheetIndex}, Row: {rowIndex}, Column: {colIndex}, Table: {tableName}, Column: {columnName}");
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Load the template workbook that contains smart markers
            Workbook workbook = new Workbook("template.xlsx");

            // Create a WorkbookDesigner and assign the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook,
                // Assign the callback to receive processing notifications
                CallBack = new SmartMarkerNotifier()
            };

            // Example JSON data source matching the smart markers in the template
            string jsonData = @"{
                'Employees': [
                    { 'Name': 'John Doe', 'Age': 30, 'Department': 'Sales' },
                    { 'Name': 'Jane Smith', 'Age': 28, 'Department': 'Marketing' }
                ]
            }";

            // Set the JSON data source with a name that matches the smart marker prefix
            designer.SetJsonDataSource("Employees", jsonData);

            // Process all smart markers in the workbook (true = preserve unrecognized markers)
            designer.Process(true);

            // Save the resulting workbook
            workbook.Save("output.xlsx");
        }
    }
}