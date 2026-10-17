// Title: Compare custom document properties of two Excel workbooks and generate an audit report with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# compare custom document properties between two .xlsx files | Generate a text file audit of differing custom properties in Excel workbooks using Aspose.Cells | How to list missing or mismatched custom properties when comparing two workbooks in .NET | C# code to compare workbook custom properties and export differences with Aspose.Cells | Case‑insensitive custom property comparison for Excel files using Aspose.Cells API
// Tags: custom document property comparison Aspose.Cells | audit report generation Excel workbooks C# | compare workbook custom properties .xlsx | case-insensitive property name matching Aspose.Cells | export property differences to text file C#

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Cells;

namespace CustomPropertyComparison
{
    // The example loads two Excel workbooks with Aspose.Cells, extracts their CustomDocumentProperties into case‑insensitive dictionaries, compares property names and values, records any missing or mismatched entries, prints the differences, and writes them to a text file for auditing.
    class Program
    {
        static void Main()
        {
            // Paths to the workbooks to compare
            string workbookPath1 = @"C:\Data\Workbook1.xlsx";
            string workbookPath2 = @"C:\Data\Workbook2.xlsx";

            // Verify that the input files exist
            if (!File.Exists(workbookPath1))
            {
                Console.WriteLine($"File not found: {workbookPath1}");
                return;
            }

            if (!File.Exists(workbookPath2))
            {
                Console.WriteLine($"File not found: {workbookPath2}");
                return;
            }

            Workbook wb1 = null;
            Workbook wb2 = null;

            try
            {
                // Load the workbooks
                wb1 = new Workbook(workbookPath1);
                wb2 = new Workbook(workbookPath2);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading workbooks: {ex.Message}");
                return;
            }

            // Retrieve custom document properties collections
            var props1 = wb1.CustomDocumentProperties;
            var props2 = wb2.CustomDocumentProperties;

            // Prepare a list to hold difference descriptions
            List<string> differences = new List<string>();

            try
            {
                // Build dictionaries for quick lookup (property name -> value)
                var dict1 = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var prop in props1)
                {
                    dict1[prop.Name] = prop.Value?.ToString() ?? string.Empty;
                }

                var dict2 = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var prop in props2)
                {
                    dict2[prop.Name] = prop.Value?.ToString() ?? string.Empty;
                }

                // Find properties present in workbook1 but missing or different in workbook2
                foreach (var kvp in dict1)
                {
                    if (!dict2.TryGetValue(kvp.Key, out string valueInWb2))
                    {
                        differences.Add($"Property '{kvp.Key}' exists in Workbook1 but is missing in Workbook2.");
                    }
                    else if (!string.Equals(kvp.Value, valueInWb2, StringComparison.Ordinal))
                    {
                        differences.Add($"Property '{kvp.Key}' differs: Workbook1='{kvp.Value}' vs Workbook2='{valueInWb2}'.");
                    }
                }

                // Find properties present in workbook2 but missing in workbook1
                foreach (var kvp in dict2)
                {
                    if (!dict1.ContainsKey(kvp.Key))
                    {
                        differences.Add($"Property '{kvp.Key}' exists in Workbook2 but is missing in Workbook1.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing custom properties: {ex.Message}");
                return;
            }

            // Output the differences
            if (differences.Count == 0)
            {
                Console.WriteLine("No differences found in custom property sets.");
            }
            else
            {
                Console.WriteLine("Differences in custom property sets:");
                foreach (var diff in differences)
                {
                    Console.WriteLine("- " + diff);
                }
            }

            // Optionally, write the differences to a text file for auditing
            string auditReportPath = @"C:\Data\CustomPropertyDifferences.txt";

            try
            {
                // Ensure the directory exists
                string dir = Path.GetDirectoryName(auditReportPath);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                File.WriteAllLines(auditReportPath, differences);
                Console.WriteLine($"Audit report saved to: {auditReportPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error writing audit report: {ex.Message}");
            }
        }
    }
}
