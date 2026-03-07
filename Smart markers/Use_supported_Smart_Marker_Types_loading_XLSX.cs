using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace SmartMarkerDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the template workbook that contains smart markers
            string templatePath = "template.xlsx";

            // Load the template workbook (XLSX format) using default load options
            Workbook templateWorkbook = new Workbook(templatePath);

            // Initialize WorkbookDesigner with the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = templateWorkbook
            };

            // Retrieve and display all smart markers found in the template
            string[] markers = designer.GetSmartMarkers();
            Console.WriteLine("Smart markers detected in the template:");
            foreach (string marker in markers)
            {
                Console.WriteLine(marker);
            }

            // Prepare a DataTable as the data source for the smart markers
            DataTable data = new DataTable("Employees");
            data.Columns.Add("Name", typeof(string));
            data.Columns.Add("Age", typeof(int));
            data.Columns.Add("Department", typeof(string));

            data.Rows.Add("John Doe", 30, "Sales");
            data.Rows.Add("Jane Smith", 28, "Marketing");
            data.Rows.Add("Bob Johnson", 45, "Engineering");

            // Bind the DataTable to a name that matches the smart marker table name
            // Example smart marker in the template: &=Employees.Name, &=Employees.Age, etc.
            designer.SetDataSource(data);

            // Process the smart markers and populate the workbook with data
            designer.Process();

            // Save the resulting workbook
            string outputPath = "output.xlsx";
            designer.Workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Processed workbook saved to '{outputPath}'.");
        }
    }
}