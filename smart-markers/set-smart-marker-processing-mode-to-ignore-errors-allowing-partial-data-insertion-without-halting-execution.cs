// Title: Set WorkbookDesigner.SmartMarkerProcessingMode to IgnoreErrors in C# to allow partial data insertion with Aspose.Cells smart markers
// AI Prompts: Generate C# code that assigns WorkbookDesigner.SmartMarkerProcessingMode = SmartMarkerProcessingMode.IgnoreErrors before calling Process() so rows with DBNull are skipped. | Adapt an existing Aspose.Cells smart‑marker example to continue processing when the data source contains missing values and still save the workbook.
// Common Searches: Aspose.Cells C# ignore smart marker errors when data contains null | How to set SmartMarkerProcessingMode to IgnoreErrors in .NET | Process Aspose.Cells smart markers with incomplete DataTable | Partial data insertion using Aspose.Cells smart markers without throwing exception | C# skip DBNull values in Aspose.Cells smart marker template
// Tags: smart marker processing mode ignore errors | WorkbookDesigner error handling Aspose.Cells | Aspose.Cells partial data insertion C# | skip DBNull values smart markers | C# smart marker ignore errors

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerExample
{
    // The example loads a smart‑marker template, creates a DataTable that includes a DBNull value, sets WorkbookDesigner.SmartMarkerProcessingMode to IgnoreErrors, processes the markers, and saves the workbook while gracefully handling missing data.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the template workbook containing smart markers
                string templatePath = @"C:\Templates\SmartMarkerTemplate.xlsx";

                // Verify that the template file exists
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file not found: {templatePath}");
                    return;
                }

                // Load the template workbook
                Workbook workbook = new Workbook(templatePath);

                // Initialize WorkbookDesigner for Smart Marker processing
                WorkbookDesigner designer = new WorkbookDesigner(workbook);

                // Prepare a data source with partial data
                DataTable dt = new DataTable();
                dt.Columns.Add("Name", typeof(string));
                dt.Columns.Add("Score", typeof(int));

                dt.Rows.Add("Alice", 85);
                dt.Rows.Add("Bob", DBNull.Value); // Missing score – will be ignored
                dt.Rows.Add("Charlie", 92);

                // Assign the data source and process smart markers
                designer.SetDataSource(dt);
                designer.Process();

                // Ensure the output directory exists
                string outputPath = @"C:\Output\SmartMarkerResult.xlsx";
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the resulting workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
