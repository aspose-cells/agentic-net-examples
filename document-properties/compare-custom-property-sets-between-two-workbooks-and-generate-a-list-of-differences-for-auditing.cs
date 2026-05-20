using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Properties;

class Program
{
    static void Main()
    {
        // Paths to the two workbooks to compare
        string workbookPath1 = "Workbook1.xlsx";
        string workbookPath2 = "Workbook2.xlsx";

        // Load the workbooks (uses the Workbook(string) constructor)
        Workbook wb1 = new Workbook(workbookPath1);
        Workbook wb2 = new Workbook(workbookPath2);

        // Retrieve the custom document property collections
        CustomDocumentPropertyCollection props1 = wb1.CustomDocumentProperties;
        CustomDocumentPropertyCollection props2 = wb2.CustomDocumentProperties;

        // Prepare a list to hold audit messages
        List<string> differences = new List<string>();

        // Helper: create a dictionary for quick lookup of properties by name
        var dict2 = new Dictionary<string, DocumentProperty>(StringComparer.OrdinalIgnoreCase);
        foreach (DocumentProperty p in props2)
        {
            dict2[p.Name] = p;
        }

        // Check for removed or modified properties
        foreach (DocumentProperty p1 in props1)
        {
            if (!dict2.TryGetValue(p1.Name, out DocumentProperty p2))
            {
                differences.Add($"Removed: '{p1.Name}' with value '{p1.Value}'");
            }
            else
            {
                // Compare values (null‑safe)
                string val1 = p1.Value?.ToString() ?? string.Empty;
                string val2 = p2.Value?.ToString() ?? string.Empty;
                if (!val1.Equals(val2, StringComparison.Ordinal))
                {
                    differences.Add($"Modified: '{p1.Name}' changed from '{val1}' to '{val2}'");
                }
                // Remove the entry so that later we can detect additions
                dict2.Remove(p1.Name);
            }
        }

        // Remaining entries in dict2 are properties that exist only in workbook2 (added)
        foreach (var kvp in dict2)
        {
            DocumentProperty p = kvp.Value;
            differences.Add($"Added: '{p.Name}' with value '{p.Value}'");
        }

        // Output the differences to console
        Console.WriteLine("Custom Document Property Differences:");
        foreach (string diff in differences)
        {
            Console.WriteLine(diff);
        }

        // Optionally, write the audit report to a text file
        string reportPath = "CustomPropertiesAudit.txt";
        File.WriteAllLines(reportPath, differences);
        Console.WriteLine($"Audit report saved to: {reportPath}");
    }
}