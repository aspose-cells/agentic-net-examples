// Title: Compare Excel custom document properties with Aspose.Cells (.NET) and generate a differences report
// Description: Loads two Excel files, extracts their CustomDocumentProperty collections, and uses case‑insensitive dictionaries to identify added, removed, or modified properties. The differences are printed to the console and written to a new workbook (two‑column table: Difference Type, Details) for auditing purposes.
// Keywords: Aspose.Cells C# compare custom document properties | Excel custom property audit .NET | list added removed modified properties | generate property differences report | custom document property comparison | Aspose.Cells workbook comparison | Excel metadata audit | C# export differences workbook
// Common Searches: compare custom document properties Aspose.Cells | audit Excel custom properties .NET | list differences between two workbooks custom properties | export custom property changes to Excel | C# code to compare workbook metadata
// Developer Intent: Detect added, removed, or changed custom document properties between two Excel workbooks and produce an Excel report of the findings.
// Use Cases: Validate that critical metadata (e.g., Author, ReviewDate) remains unchanged after a data‑processing pipeline. | Create a change‑log for regulatory compliance by listing all custom property modifications between quarterly report versions. | Support version‑control audits by comparing production and staging workbooks and exporting the discrepancies. | Automate quality checks in CI/CD pipelines to ensure custom properties meet corporate standards before deployment.
// AI Prompts: Write C# code using Aspose.Cells to compare the CustomDocumentProperties of two Excel workbooks and output added, removed, and modified entries to a new workbook. | Provide a reusable method that returns a list of difference strings for custom properties between two Workbook objects, handling null values and case‑insensitive names. | Explain how to extend the comparison to include built‑in document properties and export the combined results as a CSV file.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsCustomPropertyComparison
{
    // Loads two Excel files, extracts their CustomDocumentProperty collections, and uses case‑insensitive dictionaries to identify added, removed, or modified properties. The differences are printed to the console and written to a new workbook (two‑column table: Difference Type, Details) for auditing purposes.
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
                {
                    Console.WriteLine($"File not found: {workbookPath1}");
                    return;
                }

                if (!File.Exists(workbookPath2))
                {
                    Console.WriteLine($"File not found: {workbookPath2}");
                    return;
                }

                // Load the workbooks
                Workbook wb1 = new Workbook(workbookPath1);
                Workbook wb2 = new Workbook(workbookPath2);

                // Retrieve the custom document property collections
                CustomDocumentPropertyCollection props1 = wb1.CustomDocumentProperties;
                CustomDocumentPropertyCollection props2 = wb2.CustomDocumentProperties;

                // Build dictionaries for fast lookup (name -> DocumentProperty)
                var dict1 = new Dictionary<string, DocumentProperty>(StringComparer.OrdinalIgnoreCase);
                foreach (DocumentProperty prop in props1)
                    dict1[prop.Name] = prop;

                var dict2 = new Dictionary<string, DocumentProperty>(StringComparer.OrdinalIgnoreCase);
                foreach (DocumentProperty prop in props2)
                    dict2[prop.Name] = prop;

                // List to hold difference descriptions
                List<string> differences = new List<string>();

                // Detect removed or modified properties
                foreach (var kvp in dict1)
                {
                    string name = kvp.Key;
                    DocumentProperty prop1 = kvp.Value;

                    if (!dict2.ContainsKey(name))
                    {
                        differences.Add($"Removed: {name} = {prop1.Value}");
                    }
                    else
                    {
                        DocumentProperty prop2 = dict2[name];
                        object val1 = prop1.Value;
                        object val2 = prop2.Value;
                        bool areEqual = (val1 == null && val2 == null) ||
                                        (val1 != null && val1.Equals(val2));

                        if (!areEqual)
                        {
                            differences.Add($"Modified: {name} from '{val1}' to '{val2}'");
                        }
                    }
                }

                // Detect added properties
                foreach (var kvp in dict2)
                {
                    string name = kvp.Key;
                    if (!dict1.ContainsKey(name))
                    {
                        DocumentProperty prop = kvp.Value;
                        differences.Add($"Added: {name} = {prop.Value}");
                    }
                }

                // Output differences to console
                Console.WriteLine("Custom Document Property Differences:");
                foreach (string diff in differences)
                    Console.WriteLine(diff);

                // Create a simple report workbook to store the differences
                Workbook reportWb = new Workbook(); // default constructor
                Worksheet sheet = reportWb.Worksheets[0];

                // Write header
                sheet.Cells[0, 0].PutValue("Difference Type");
                sheet.Cells[0, 1].PutValue("Details");

                // Populate rows
                for (int i = 0; i < differences.Count; i++)
                {
                    string diff = differences[i];
                    // Split the first word as type (Added/Removed/Modified)
                    string type = diff.Split(':')[0];
                    sheet.Cells[i + 1, 0].PutValue(type);
                    sheet.Cells[i + 1, 1].PutValue(diff);
                }

                // Save the report workbook
                string reportPath = "CustomPropertiesDifferencesReport.xlsx";
                reportWb.Save(reportPath);
                Console.WriteLine($"Report saved to {reportPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
