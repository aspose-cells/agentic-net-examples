using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsCustomPropertyAudit
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Paths to the two workbooks to compare
                string workbookPath1 = "Workbook1.xlsx";
                string workbookPath2 = "Workbook2.xlsx";

                // Verify that the input files exist
                if (!File.Exists(workbookPath1))
                    throw new FileNotFoundException($"File not found: {workbookPath1}");
                if (!File.Exists(workbookPath2))
                    throw new FileNotFoundException($"File not found: {workbookPath2}");

                // Load the workbooks
                Workbook wb1 = new Workbook(workbookPath1);
                Workbook wb2 = new Workbook(workbookPath2);

                // Retrieve custom document properties collections
                CustomDocumentPropertyCollection props1 = wb1.CustomDocumentProperties;
                CustomDocumentPropertyCollection props2 = wb2.CustomDocumentProperties;

                // Build dictionaries for fast lookup (property name -> DocumentProperty)
                var dict1 = new Dictionary<string, DocumentProperty>(StringComparer.OrdinalIgnoreCase);
                foreach (DocumentProperty prop in props1)
                    dict1[prop.Name] = prop;

                var dict2 = new Dictionary<string, DocumentProperty>(StringComparer.OrdinalIgnoreCase);
                foreach (DocumentProperty prop in props2)
                    dict2[prop.Name] = prop;

                // List to hold difference descriptions
                List<string> differences = new List<string>();

                // Detect removed or modified properties (present in wb1 but not or changed in wb2)
                foreach (var kvp in dict1)
                {
                    string name = kvp.Key;
                    DocumentProperty prop1 = kvp.Value;

                    if (!dict2.TryGetValue(name, out DocumentProperty prop2))
                    {
                        differences.Add($"Removed: {name} = {prop1.Value}");
                    }
                    else
                    {
                        // Compare values (using string representation for simplicity)
                        string val1 = prop1.Value?.ToString() ?? string.Empty;
                        string val2 = prop2.Value?.ToString() ?? string.Empty;

                        if (!val1.Equals(val2, StringComparison.Ordinal))
                        {
                            differences.Add($"Modified: {name} from '{val1}' to '{val2}'");
                        }
                    }
                }

                // Detect added properties (present in wb2 but not in wb1)
                foreach (var kvp in dict2)
                {
                    string name = kvp.Key;
                    if (!dict1.ContainsKey(name))
                    {
                        DocumentProperty prop2 = kvp.Value;
                        differences.Add($"Added: {name} = {prop2.Value}");
                    }
                }

                // Create a new workbook to store the audit report
                Workbook reportWorkbook = new Workbook();

                // Add a worksheet for the differences and obtain the worksheet reference
                Worksheet sheet = reportWorkbook.Worksheets.Add("Differences");

                // Write header
                sheet.Cells["A1"].PutValue("Custom Property Differences");
                sheet.Cells["A2"].PutValue("Type");
                sheet.Cells["B2"].PutValue("Property Name");
                sheet.Cells["C2"].PutValue("Details");

                // Populate the sheet with differences
                int row = 3;
                foreach (string diff in differences)
                {
                    // Expected format: "Removed: Name = Value", "Modified: Name from 'old' to 'new'", etc.
                    int colonPos = diff.IndexOf(':');
                    if (colonPos > -1)
                    {
                        string type = diff.Substring(0, colonPos).Trim();
                        string rest = diff.Substring(colonPos + 1).Trim();

                        sheet.Cells[row, 0].PutValue(type); // Column A

                        // Extract property name (up to first '=' or space)
                        int eqPos = rest.IndexOf('=');
                        string propName = eqPos > -1 ? rest.Substring(0, eqPos).Trim() : rest;
                        sheet.Cells[row, 1].PutValue(propName); // Column B

                        sheet.Cells[row, 2].PutValue(rest); // Column C (full detail)
                    }
                    else
                    {
                        // Fallback: write whole line in details column
                        sheet.Cells[row, 2].PutValue(diff);
                    }
                    row++;
                }

                // Auto-fit columns for better readability
                sheet.AutoFitColumns();

                // Save the report workbook
                string reportPath = "CustomPropertiesAuditReport.xlsx";
                reportWorkbook.Save(reportPath);

                // Output differences to console for immediate feedback
                Console.WriteLine("Custom Property Differences:");
                foreach (string diff in differences)
                {
                    Console.WriteLine(diff);
                }

                Console.WriteLine($"Audit report saved to: {reportPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}