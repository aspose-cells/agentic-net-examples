// Title: Create a .NET program that scans .xlsx workbooks and reports optional ContentTypeProperties missing the IsNillable flag using Aspose.Cells
// Common Searches: how to list optional ContentTypeProperties that are not nillable in Excel files with Aspose.Cells C# | C# batch process .xlsx files and export properties where IsOptional is true and IsNillable is false | using reflection to access the ContentTypeProperties collection in an Aspose.Cells workbook | generate a summary workbook of missing IsNillable flags across multiple Excel documents | Aspose.Cells example for reporting custom content type properties without IsNillable attribute
// Tags: enumerate workbook ContentTypeProperties via reflection | detect missing IsNillable flag in optional properties | create summary Excel report of property validation results | batch scan .xlsx files with Aspose.Cells for property analysis | export ContentTypeProperties details to a new workbook

using System;
using System.IO;
using System.Collections;
using Aspose.Cells;

namespace ContentTypePropertyReport
{
    // The C# utility iterates through all .xlsx files in a folder, loads each workbook with Aspose.Cells, uses reflection to read the ContentTypeProperties collection, identifies properties where IsOptional is true and IsNillable is false, and writes the workbook name, property name, value, and flag states into a new Excel report saved to a specified location.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing the workbooks to analyze
            string inputFolder = @"C:\Workbooks";

            // Path for the generated report workbook
            string reportPath = @"C:\Report\OptionalPropertiesWithoutIsNillable.xlsx";

            // Ensure input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            // Ensure output directory exists
            string reportDir = Path.GetDirectoryName(reportPath);
            if (!Directory.Exists(reportDir))
                Directory.CreateDirectory(reportDir);

            // Create a new workbook for the report
            Workbook reportWorkbook = new Workbook();
            Worksheet reportSheet = reportWorkbook.Worksheets[0];
            reportSheet.Name = "Report";

            // Write header row
            reportSheet.Cells[0, 0].PutValue("Workbook");
            reportSheet.Cells[0, 1].PutValue("Property Name");
            reportSheet.Cells[0, 2].PutValue("Property Value");
            reportSheet.Cells[0, 3].PutValue("IsOptional");
            reportSheet.Cells[0, 4].PutValue("IsNillable");

            int reportRow = 1; // Start writing data from the second row

            // Iterate through all Excel files in the input folder
            foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx"))
            {
                if (!File.Exists(filePath))
                    continue; // Safety check

                try
                {
                    // Load the workbook
                    Workbook wb = new Workbook(filePath);

                    // Use reflection to access ContentTypeProperties without a direct type reference
                    var propCollectionObj = wb.GetType()
                                             .GetProperty("ContentTypeProperties")
                                             ?.GetValue(wb);

                    if (propCollectionObj is IEnumerable propEnumerable)
                    {
                        foreach (var propObj in propEnumerable)
                        {
                            if (propObj == null) continue;

                            var propType = propObj.GetType();

                            // Retrieve required property values via reflection
                            bool isOptional = Convert.ToBoolean(propType.GetProperty("IsOptional")?.GetValue(propObj) ?? false);
                            bool isNillable = Convert.ToBoolean(propType.GetProperty("IsNillable")?.GetValue(propObj) ?? false);

                            // Check condition: optional and NOT nillable
                            if (isOptional && !isNillable)
                            {
                                string name = propType.GetProperty("Name")?.GetValue(propObj)?.ToString() ?? string.Empty;
                                string value = propType.GetProperty("Value")?.GetValue(propObj)?.ToString() ?? string.Empty;

                                // Write details to the report sheet
                                reportSheet.Cells[reportRow, 0].PutValue(Path.GetFileName(filePath));
                                reportSheet.Cells[reportRow, 1].PutValue(name);
                                reportSheet.Cells[reportRow, 2].PutValue(value);
                                reportSheet.Cells[reportRow, 3].PutValue(isOptional);
                                reportSheet.Cells[reportRow, 4].PutValue(isNillable);
                                reportRow++;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                    // Continue with next file
                }
            }

            // Auto-fit columns for better readability
            reportSheet.AutoFitColumns();

            try
            {
                // Save the report workbook
                reportWorkbook.Save(reportPath, SaveFormat.Xlsx);
                Console.WriteLine($"Report saved to: {reportPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save report: {ex.Message}");
            }
        }
    }
}
