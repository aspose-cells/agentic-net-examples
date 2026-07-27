using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerErrorHandling
{
    // Callback to log each smart marker being processed.
    // This is optional but helps to identify which marker caused the issue.
    public class SmartMarkerLogger : ISmartMarkerCallBack
    {
        public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
        {
            Console.WriteLine($"Processing marker - Sheet:{sheetIndex}, Row:{rowIndex}, Column:{colIndex}, Table:{tableName}, Column:{columnName}");
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Load the template workbook that contains smart markers.
            // The template should have a marker like "&=$Employees.Name" and "&=$Employees.Salary".
            Workbook template = new Workbook("SmartMarkerTemplate.xlsx");

            // Prepare a data source that intentionally lacks the "Salary" column.
            DataTable employees = new DataTable("Employees");
            employees.Columns.Add("Name", typeof(string));
            // Note: "Salary" column is missing to simulate the error condition.
            employees.Rows.Add("John Doe");
            employees.Rows.Add("Jane Smith");

            // Set up the WorkbookDesigner.
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = template,
                CallBack = new SmartMarkerLogger() // optional logging
            };

            // Bind the incomplete data source.
            designer.SetDataSource(employees);

            // Process the smart markers with error handling.
            try
            {
                // The boolean parameter indicates whether unrecognized smart markers are preserved.
                // Setting it to false will cause an exception if a marker cannot be resolved.
                designer.Process(false);
                Console.WriteLine("Smart markers processed successfully.");
            }
            catch (Exception ex)
            {
                // Catch exceptions caused by missing fields in the data source.
                Console.WriteLine("Error processing smart markers: " + ex.Message);
                // Additional handling such as logging or fallback logic can be placed here.
            }

            // Save the resulting workbook.
            designer.Workbook.Save("SmartMarkerResult.xlsx");
        }
    }
}