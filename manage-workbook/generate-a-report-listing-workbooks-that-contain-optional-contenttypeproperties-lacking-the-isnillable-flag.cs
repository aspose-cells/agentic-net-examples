// Title: Create a C# program with Aspose.Cells to report Excel workbooks that have optional ContentTypeProperties missing the IsNillable flag
// Common Searches: C# Aspose.Cells find Excel workbooks where optional ContentTypeProperties are not nillable | how to generate a list of .xlsx files with missing IsNillable flag in ContentTypeProperties using Aspose.Cells | scan directory for Excel files with optional non‑nillable ContentTypeProperties in .NET | Aspose.Cells reflection example to check IsOptional and IsNillable properties in workbooks | report workbooks containing optional ContentTypeProperties without IsNillable flag in C#
// Tags: Aspose.Cells optional ContentTypeProperties detection | C# reflection inspection of ContentTypeProperties | Excel workbook scanning for missing IsNillable flag | batch processing .xlsx files with Aspose.Cells | generate compliance report for non‑nillable optional properties

using System;
using System.IO;
using System.Collections.Generic;
using System.Reflection;
using Aspose.Cells;

namespace ContentTypePropertyReport
{
    // The C# utility recursively enumerates .xlsx files in a folder, loads each workbook with Aspose.Cells, uses reflection to read the IsOptional and IsNillable attributes of every ContentTypeProperty, records workbooks that contain an optional property where IsNillable is false, and writes the full paths and a total count to a text report.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing the workbooks to scan
            string workbooksFolder = @"C:\Workbooks";

            // Output report file
            string reportPath = @"C:\Report\WorkbooksWithMissingIsNillable.txt";

            // Ensure the output directory exists
            string reportDir = Path.GetDirectoryName(reportPath);
            if (!string.IsNullOrEmpty(reportDir))
            {
                Directory.CreateDirectory(reportDir);
            }

            // List to hold workbook names that meet the criteria
            List<string> workbooksWithIssues = new List<string>();

            // Get all Excel files in the specified folder (including subfolders)
            string[] excelFiles = Directory.GetFiles(workbooksFolder, "*.xlsx", SearchOption.AllDirectories);

            foreach (string filePath in excelFiles)
            {
                // Verify the file exists before attempting to load
                if (!File.Exists(filePath))
                {
                    continue;
                }

                try
                {
                    // Load the workbook using Aspose.Cells
                    Workbook workbook = new Workbook(filePath);

                    // Flag to indicate if this workbook has at least one offending property
                    bool hasIssue = false;

                    // Iterate through all ContentTypeProperties of the workbook
                    foreach (object propObj in workbook.ContentTypeProperties)
                    {
                        // Use reflection to access IsOptional and IsNillable properties
                        Type propType = propObj.GetType();
                        PropertyInfo optionalProp = propType.GetProperty("IsOptional");
                        PropertyInfo nillableProp = propType.GetProperty("IsNillable");

                        if (optionalProp != null && nillableProp != null)
                        {
                            bool isOptional = Convert.ToBoolean(optionalProp.GetValue(propObj));
                            bool isNillable = Convert.ToBoolean(nillableProp.GetValue(propObj));

                            // Check for optional properties that are NOT nillable
                            if (isOptional && !isNillable)
                            {
                                hasIssue = true;
                                break; // No need to continue scanning this workbook
                            }
                        }
                    }

                    if (hasIssue)
                    {
                        // Add the workbook full path to the report list
                        workbooksWithIssues.Add(filePath);
                    }
                }
                catch (Exception ex)
                {
                    // Log or handle errors for individual workbooks without stopping the whole process
                    Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
                }
            }

            // Write the report
            try
            {
                using (StreamWriter writer = new StreamWriter(reportPath))
                {
                    writer.WriteLine("Workbooks containing optional ContentTypeProperties lacking the IsNillable flag:");
                    writer.WriteLine("--------------------------------------------------------------------------");
                    foreach (string wb in workbooksWithIssues)
                    {
                        writer.WriteLine(wb);
                    }
                    writer.WriteLine("--------------------------------------------------------------------------");
                    writer.WriteLine($"Total workbooks found: {workbooksWithIssues.Count}");
                }

                Console.WriteLine("Report generation completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to write report: {ex.Message}");
            }
        }
    }
}
