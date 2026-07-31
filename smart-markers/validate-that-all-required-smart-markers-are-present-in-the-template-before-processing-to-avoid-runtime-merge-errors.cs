// Title: C# – Validate Required Smart Markers in Aspose.Cells Template to Prevent Merge Errors
// Description: C# sample that loads or creates an Excel template, extracts smart markers with WorkbookDesigner.GetSmartMarkers(), compares them against a required list, reports any missing markers, throws an exception, then binds a matching DataTable and processes the merge. Guarantees template integrity before smart‑marker processing.
// Keywords: Aspose.Cells | C# | smart markers | WorkbookDesigner | GetSmartMarkers | template validation | missing smart markers | Excel merge error | data export | smart marker validation | code example | GitHub
// Common Searches: aspocells validate smart markers c# | check smart markers in excel template | WorkbookDesigner GetSmartMarkers usage | throw exception when smart markers missing | create excel template with required smart markers
// Developer Intent: Ensure that every required smart marker is present in the workbook before running the smart‑marker merge.
// Use Cases: Pre‑flight validation of reporting templates to avoid runtime merge failures. | Automated creation of a minimal workbook when the expected template is absent. | Logging and aborting a data‑export pipeline when required markers are missing, preserving data integrity.
// AI Prompts: Write C# code using Aspose.Cells that checks a list of required smart markers in a workbook and throws a detailed exception for any that are absent. | Show how to programmatically generate an Excel template containing specific smart markers with Aspose.Cells WorkbookDesigner. | Suggest a more efficient, case‑insensitive method to validate required smart markers in an Aspose.Cells template.

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace SmartMarkerValidationDemo
{
    // C# sample that loads or creates an Excel template, extracts smart markers with WorkbookDesigner.GetSmartMarkers(), compares them against a required list, reports any missing markers, throws an exception, then binds a matching DataTable and processes the merge. Guarantees template integrity before smart‑marker processing.
    class Program
    {
        static void Main()
        {
            const string templatePath = "template.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Ensure the template file exists; if not, create a simple one with required smart markers
                if (!File.Exists(templatePath))
                {
                    CreateTemplateWorkbook(templatePath);
                }

                // Load the template workbook
                Workbook templateWorkbook = new Workbook(templatePath);

                // Initialize the WorkbookDesigner with the loaded workbook
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = templateWorkbook
                };

                // Retrieve all smart markers present in the template
                string[] existingMarkers = designer.GetSmartMarkers();

                // Define the list of smart markers that must be present in the template
                string[] requiredMarkers = new string[]
                {
                    "&=COLORS_TIMES.COLORS",
                    "&=COLORS_TIMES.TIMES"
                };

                // Validate required markers
                List<string> missingMarkers = new List<string>();
                foreach (string required in requiredMarkers)
                {
                    bool found = false;
                    foreach (string existing in existingMarkers)
                    {
                        if (string.Equals(existing, required, StringComparison.OrdinalIgnoreCase))
                        {
                            found = true;
                            break;
                        }
                    }
                    if (!found)
                    {
                        missingMarkers.Add(required);
                    }
                }

                if (missingMarkers.Count > 0)
                {
                    Console.WriteLine("The following required smart markers are missing from the template:");
                    foreach (string marker in missingMarkers)
                    {
                        Console.WriteLine($"  {marker}");
                    }
                    throw new InvalidOperationException("Template validation failed due to missing smart markers.");
                }

                // Sample data source matching the smart markers
                DataTable data = new DataTable("COLORS_TIMES");
                data.Columns.Add("COLORS", typeof(string));
                data.Columns.Add("TIMES", typeof(DateTime));
                data.Rows.Add("Red", DateTime.Now.AddDays(-1));
                data.Rows.Add("Yellow", DateTime.Now);
                data.Rows.Add("Green", DateTime.Now.AddDays(1));

                // Bind the data source and process the smart markers
                designer.SetDataSource(data);
                designer.Process();

                // Save the processed workbook
                designer.Workbook.Save(outputPath);
                Console.WriteLine($"Processing completed successfully. Output saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Creates a minimal template workbook containing the required smart markers
        private static void CreateTemplateWorkbook(string path)
        {
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            ws.Name = "Data";

            // Place smart markers in the first row
            ws.Cells["A1"].PutValue("&=COLORS_TIMES.COLORS");
            ws.Cells["B1"].PutValue("&=COLORS_TIMES.TIMES");

            wb.Save(path);
            Console.WriteLine($"Template workbook created at '{path}'.");
        }
    }
}
