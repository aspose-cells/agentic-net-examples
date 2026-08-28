// Title: Processing identical smart markers on every worksheet with WorkbookDesigner.Process in Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel template, binds a DataTable to smart markers, loops through all worksheet indices, and calls WorkbookDesigner.Process(index, true) to preserve unrecognized markers. | Show how to use Aspose.Cells WorkbookDesigner to apply the same smart‑marker data source to multiple sheets and save the resulting workbook. | Provide an example of iterating over workbook.Worksheets.Count and processing smart markers on each sheet while keeping unknown markers intact.
// Common Searches: Aspose.Cells C# process smart markers on each sheet of a workbook | WorkbookDesigner.Process multiple worksheets example | preserve unknown smart markers when using Aspose.Cells | bind DataTable to smart markers across several worksheets Aspose.Cells | save processed workbook after smart marker iteration .NET
// Tags: process smart markers per worksheet Aspose.Cells | WorkbookDesigner.Process preserve unknown markers | bind DataTable to smart markers Aspose.Cells | iterate worksheets workbook.Worksheets.Count C# | save workbook after smart marker processing .NET | template with identical smart markers multiple sheets

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerProcessing
{
    // The example loads a template workbook that contains identical smart markers on every worksheet, creates a DataTable with employee data, sets it as the data source for a WorkbookDesigner, iterates through each worksheet index calling designer.Process(index, true) to process and preserve unknown markers, and finally saves the processed workbook to a new file.
    public class MultipleSheetProcessor
    {
        public static void Run()
        {
            try
            {
                string templatePath = "TemplateWithSmartMarkers.xlsx";

                // Ensure the template file exists before loading
                if (!File.Exists(templatePath))
                {
                    throw new FileNotFoundException($"Template file not found: {templatePath}");
                }

                // Load the template workbook that contains identical smart markers on each worksheet
                Workbook workbook = new Workbook(templatePath);

                // Prepare a sample data source (DataTable) that will be bound to the smart markers
                DataTable dataTable = new DataTable("Employees");
                dataTable.Columns.Add("Name", typeof(string));
                dataTable.Columns.Add("Age", typeof(int));
                dataTable.Columns.Add("Department", typeof(string));

                dataTable.Rows.Add("John Doe", 30, "Sales");
                dataTable.Rows.Add("Jane Smith", 28, "Marketing");
                dataTable.Rows.Add("Bob Johnson", 35, "IT");

                // Initialize the WorkbookDesigner and assign the loaded workbook
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };

                // Set the data source once; it will be used for all sheets
                designer.SetDataSource(dataTable);

                // Iterate through each worksheet and process smart markers on that sheet
                // Process(int sheetIndex, bool isPreserved) works on worksheet level
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    // true = preserve unrecognized smart markers (adjust as needed)
                    designer.Process(i, true);
                }

                // Save the processed workbook
                string outputPath = "ProcessedMultipleSheets.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook processed and saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            MultipleSheetProcessor.Run();
        }
    }
}
