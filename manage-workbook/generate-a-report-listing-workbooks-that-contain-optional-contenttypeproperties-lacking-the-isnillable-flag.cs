// Title: Generate a C# Aspose.Cells report of workbooks that have optional ContentTypeProperties missing the IsNillable flag
// AI Prompts: Write a C# console program that uses Aspose.Cells to iterate over every .xlsx file in a given folder, selects ContentTypeProperties where IsOptional is true and IsNillable is false, and writes the workbook name, property name, and flag values into a new Excel summary workbook. | Extend the program to also capture each property's data type, sort the rows by workbook name, and add a timestamp column indicating when the report entry was created. | Add comprehensive error handling that logs unreadable or corrupted workbooks to a separate text file, including the exception message and file path, while allowing the scan to continue.
// Common Searches: aspocells c# list optional contenttypeproperties without isnillable | how to create a report of missing isnillable flag in excel files using Aspose.Cells | scan a directory for contenttypeproperty attributes in .xlsx files c# | export optional contenttypeproperty details to a new workbook with Aspose.Cells
// Tags: list optional ContentTypeProperties Aspose.Cells | detect missing IsNillable flag Excel | C# folder scan workbooks Aspose.Cells | export ContentTypeProperty analysis to Excel | dynamic property access Aspose.Cells

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace WorkbookContentTypeReport
{
    // Scans a specified folder for .xlsx files, identifies ContentTypeProperties that are optional but not nillable, and writes workbook name, property name, and flag values to a generated Excel report using Aspose.Cells, with optional logging for errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing the workbooks to be scanned
            string inputFolder = @"C:\Workbooks";

            // Path for the generated report workbook
            string reportPath = @"C:\Report\OptionalContentTypePropertiesReport.xlsx";

            // Ensure the input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            // Ensure the directory for the report exists
            string reportDir = Path.GetDirectoryName(reportPath);
            if (!Directory.Exists(reportDir))
            {
                Directory.CreateDirectory(reportDir);
            }

            // DataTable to collect report data
            DataTable reportTable = new DataTable();
            reportTable.Columns.Add("Workbook", typeof(string));
            reportTable.Columns.Add("ContentTypeProperty", typeof(string));
            reportTable.Columns.Add("IsOptional", typeof(bool));
            reportTable.Columns.Add("IsNillable", typeof(bool));

            // Iterate through all Excel files in the input folder
            foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx"))
            {
                try
                {
                    // Verify the file exists before loading
                    if (!File.Exists(filePath))
                        continue;

                    // Load the workbook using Aspose.Cells
                    Workbook wb = new Workbook(filePath);

                    // Access the collection of ContentTypeProperties
                    var contentTypeProperties = wb.ContentTypeProperties;

                    // Examine each ContentTypeProperty using dynamic to avoid compile‑time type dependency
                    foreach (var prop in contentTypeProperties)
                    {
                        try
                        {
                            dynamic d = prop; // Resolve at runtime
                            if (d.IsOptional && !d.IsNillable)
                            {
                                reportTable.Rows.Add(
                                    Path.GetFileName(filePath),
                                    d.Name,
                                    (bool)d.IsOptional,
                                    (bool)d.IsNillable);
                            }
                        }
                        catch (Exception ex)
                        {
                            // Log and continue with next property
                            Console.WriteLine($"Error processing property in '{filePath}': {ex.Message}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Log and continue with next workbook
                    Console.WriteLine($"Error processing workbook '{filePath}': {ex.Message}");
                }
            }

            // Create a new workbook for the report
            Workbook reportWorkbook = new Workbook();
            Worksheet sheet = reportWorkbook.Worksheets[0];
            sheet.Name = "OptionalProperties";

            // Write header row
            sheet.Cells["A1"].PutValue("Workbook");
            sheet.Cells["B1"].PutValue("ContentTypeProperty");
            sheet.Cells["C1"].PutValue("IsOptional");
            sheet.Cells["D1"].PutValue("IsNillable");

            // Populate rows from the DataTable
            for (int i = 0; i < reportTable.Rows.Count; i++)
            {
                DataRow row = reportTable.Rows[i];
                int excelRow = i + 2; // +2 because of header row

                sheet.Cells[$"A{excelRow}"].PutValue(row["Workbook"]);
                sheet.Cells[$"B{excelRow}"].PutValue(row["ContentTypeProperty"]);
                sheet.Cells[$"C{excelRow}"].PutValue(row["IsOptional"]);
                sheet.Cells[$"D{excelRow}"].PutValue(row["IsNillable"]);
            }

            // Auto‑fit columns for better readability
            sheet.AutoFitColumns();

            // Save the report workbook with safety handling
            try
            {
                reportWorkbook.Save(reportPath);
                Console.WriteLine($"Report saved to: {reportPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save report: {ex.Message}");
            }
        }
    }
}
