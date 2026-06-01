using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace SmartMarkerValidationDemo
{
    public class SmartMarkerValidator
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string templatePath = "TemplateWithSmartMarkers.xlsx";
            const string outputPath = "ProcessedOutput.xlsx";

            // Verify that the template file exists to avoid FileNotFoundException
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {templatePath}");
                return;
            }

            // Load the template workbook that contains smart markers
            Workbook templateWorkbook = new Workbook(templatePath);

            // Initialize the WorkbookDesigner with the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner(templateWorkbook);

            // ----- Set up a sample data source -----
            // In a real scenario replace this with your actual data source
            DataTable dt = new DataTable("Employees");
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Age", typeof(int));
            dt.Rows.Add("John Doe", 30);
            dt.Rows.Add("Jane Smith", 28);
            designer.SetDataSource("Employees", dt);
            // ---------------------------------------

            // Process the smart markers – this should replace all markers with data
            designer.Process();

            // After processing, retrieve any remaining smart markers
            string[] unresolvedMarkers = designer.GetSmartMarkers();

            // Validate that no unresolved markers remain
            if (unresolvedMarkers.Length == 0)
            {
                Console.WriteLine("All smart markers have been successfully resolved.");
            }
            else
            {
                Console.WriteLine("Unresolved smart markers found:");
                foreach (string marker in unresolvedMarkers)
                {
                    Console.WriteLine($"- {marker}");
                }
            }

            // Save the processed workbook (optional)
            try
            {
                designer.Workbook.Save(outputPath);
                Console.WriteLine($"Processed workbook saved to: {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save the workbook: {saveEx.Message}");
            }
        }
    }
}