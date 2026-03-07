using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerVariablesDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the template workbook that contains smart markers and a variables sheet
            string templatePath = "TemplateWithVariables.xlsx";

            // If the template does not exist, create a simple one programmatically
            if (!File.Exists(templatePath))
            {
                Workbook wb = new Workbook();
                // Sheet with smart markers
                Worksheet dataSheet = wb.Worksheets[0];
                dataSheet.Name = "Data";
                dataSheet.Cells["A1"].PutValue("&=ReportDate"); // variable smart marker
                dataSheet.Cells["A3"].PutValue("&=Employees.Name"); // regular smart marker
                dataSheet.Cells["B3"].PutValue("&=Employees.Age");
                dataSheet.Cells["C3"].PutValue("&=Employees.Department");

                // Variables worksheet (can be empty)
                Worksheet varSheet = wb.Worksheets.Add("Variables");

                wb.Save(templatePath);
            }

            // Load the template workbook
            Workbook templateWorkbook = new Workbook(templatePath);

            // Initialize WorkbookDesigner with the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = templateWorkbook,
                VariablesWorksheetName = "Variables"
            };

            // OPTIONAL: List all smart markers found in the workbook
            Console.WriteLine("Smart markers detected in the template:");
            foreach (string marker in designer.GetSmartMarkers())
            {
                Console.WriteLine(marker);
            }

            // Prepare data sources
            // Simple variable replacement
            DateTime reportDate = DateTime.Today;
            designer.SetDataSource("ReportDate", reportDate);

            // Table data source for regular smart markers
            DataTable employeeTable = new DataTable("Employees");
            employeeTable.Columns.Add("Name", typeof(string));
            employeeTable.Columns.Add("Age", typeof(int));
            employeeTable.Columns.Add("Department", typeof(string));
            employeeTable.Rows.Add("John Doe", 30, "Sales");
            employeeTable.Rows.Add("Jane Smith", 28, "Marketing");
            employeeTable.Rows.Add("Bob Johnson", 35, "IT");
            designer.SetDataSource(employeeTable);

            // Process all smart markers, including variables
            designer.Process();

            // Save the populated workbook
            string outputPath = "OutputWithVariables.xlsx";
            designer.Workbook.Save(outputPath);

            Console.WriteLine($"Workbook processed and saved to '{outputPath}'.");
        }
    }
}