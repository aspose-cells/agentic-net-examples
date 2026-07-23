// Title: Audit custom document properties in two Excel workbooks with Aspose.Cells for .NET
// Description: Loads two .xlsx files, extracts each workbook's CustomDocumentPropertyCollection, and performs a case‑insensitive comparison. The code identifies added, modified, and removed custom properties, builds readable diff strings, and writes the audit list to the console.
// Keywords: Aspose.Cells C# compare custom properties | Excel custom document properties audit | detect added modified removed workbook metadata | case‑insensitive property comparison .NET | custom property differences between workbooks | C# Excel metadata comparison example | Aspose.Cells custom document property collection
// Common Searches: compare custom document properties between two Excel files using Aspose.Cells | list differences in workbook custom properties .NET | how to audit Excel custom metadata with C# | detect added or removed custom properties in Excel workbooks | Aspose.Cells example for property change log
// Developer Intent: Identify and report added, modified, or removed custom document properties across two Excel workbooks.
// Use Cases: Generate a version‑control change log that captures metadata updates between released workbook editions. | Validate that required custom properties are present after automated processing and flag any missing or altered entries. | Create compliance reports that record property additions, deletions, or value changes for regulatory audits. | Synchronize metadata across multiple workbook copies by detecting discrepancies before deployment.
// AI Prompts: Refactor CompareCustomDocumentProperties to return a list of objects with fields: Name, OldValue, NewValue, ChangeType. | Write MSTest/NUnit unit tests covering scenarios where properties are added, modified, or removed. | Show how to export the differences to a CSV or Excel sheet using Aspose.Cells after the comparison. | Add support for comparing built‑in document properties (author, title, etc.) alongside custom properties.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Properties;

// Loads two .xlsx files, extracts each workbook's CustomDocumentPropertyCollection, and performs a case‑insensitive comparison. The code identifies added, modified, and removed custom properties, builds readable diff strings, and writes the audit list to the console.
class WorkbookCustomPropertyComparer
{
    static void Main()
    {
        // Paths to the two workbooks to compare
        string workbookPath1 = "Workbook1.xlsx";
        string workbookPath2 = "Workbook2.xlsx";

        // Verify that both files exist before attempting to load them
        if (!File.Exists(workbookPath1) || !File.Exists(workbookPath2))
        {
            Console.WriteLine("One or both workbook files were not found:");
            if (!File.Exists(workbookPath1)) Console.WriteLine($" - {workbookPath1}");
            if (!File.Exists(workbookPath2)) Console.WriteLine($" - {workbookPath2}");
            return;
        }

        try
        {
            // Load the workbooks
            Workbook workbook1 = new Workbook(workbookPath1);
            Workbook workbook2 = new Workbook(workbookPath2);

            // Compare their custom document properties
            List<string> differences = CompareCustomDocumentProperties(workbook1, workbook2);

            // Output the differences
            Console.WriteLine("Custom Document Property Differences:");
            foreach (string diff in differences)
            {
                Console.WriteLine(diff);
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Compares the CustomDocumentProperties collections of two workbooks
    static List<string> CompareCustomDocumentProperties(Workbook wbA, Workbook wbB)
    {
        var diffs = new List<string>();

        // Retrieve the custom property collections
        CustomDocumentPropertyCollection propsA = wbA.CustomDocumentProperties;
        CustomDocumentPropertyCollection propsB = wbB.CustomDocumentProperties;

        // Build dictionaries for fast lookup by property name (case‑insensitive)
        var dictA = new Dictionary<string, DocumentProperty>(StringComparer.OrdinalIgnoreCase);
        foreach (DocumentProperty prop in propsA)
        {
            dictA[prop.Name] = prop;
        }

        var dictB = new Dictionary<string, DocumentProperty>(StringComparer.OrdinalIgnoreCase);
        foreach (DocumentProperty prop in propsB)
        {
            dictB[prop.Name] = prop;
        }

        // Detect added or modified properties in workbook B
        foreach (var kvp in dictB)
        {
            string name = kvp.Key;
            DocumentProperty propB = kvp.Value;

            if (!dictA.ContainsKey(name))
            {
                // Property exists only in B
                diffs.Add($"Added: {name} = {propB.Value} (Type: {propB.Type})");
            }
            else
            {
                // Property exists in both; check for value change
                DocumentProperty propA = dictA[name];
                if (!object.Equals(propA.Value, propB.Value))
                {
                    diffs.Add($"Modified: {name} from '{propA.Value}' to '{propB.Value}'");
                }
            }
        }

        // Detect removed properties (present in A but not in B)
        foreach (var kvp in dictA)
        {
            string name = kvp.Key;
            DocumentProperty propA = kvp.Value;

            if (!dictB.ContainsKey(name))
            {
                diffs.Add($"Removed: {name} = {propA.Value} (Type: {propA.Type})");
            }
        }

        return diffs;
    }
}
