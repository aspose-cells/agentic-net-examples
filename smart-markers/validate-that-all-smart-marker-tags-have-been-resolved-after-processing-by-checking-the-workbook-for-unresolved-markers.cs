using System;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerValidationDemo
{
    public class Program
    {
        public static void Main()
        {
            // Load the template workbook that contains smart markers
            Workbook templateWorkbook = new Workbook("SmartMarkerTemplate.xlsx");

            // Initialize the WorkbookDesigner with the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner(templateWorkbook);

            // -------------------------------------------------
            // Prepare a sample data source (replace with real data)
            // -------------------------------------------------
            DataTable dataTable = new DataTable("Employees");
            dataTable.Columns.Add("Name", typeof(string));
            dataTable.Columns.Add("Age", typeof(int));
            dataTable.Rows.Add("John Doe", 30);
            dataTable.Rows.Add("Jane Smith", 28);

            // Bind the data source to the designer
            designer.SetDataSource("Employees", dataTable);

            // Process the smart markers – this should replace all markers with data
            designer.Process();

            // After processing, retrieve any remaining (unresolved) smart markers
            string[] unresolvedMarkers = designer.GetSmartMarkers();

            // Check if any markers are still present
            if (unresolvedMarkers.Length > 0)
            {
                Console.WriteLine("Unresolved smart markers found:");
                foreach (string marker in unresolvedMarkers)
                {
                    Console.WriteLine($"- {marker}");
                }
            }
            else
            {
                Console.WriteLine("All smart markers have been successfully resolved.");
            }

            // Save the processed workbook
            designer.Workbook.Save("SmartMarkerResult.xlsx");
        }
    }
}