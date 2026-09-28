// Title: Generate an audit report of differing custom document properties between two Excel workbooks using Aspose.Cells for .NET
// AI Prompts: Write a C# method that takes two Aspose.Cells Workbook objects and returns a list describing any mismatched or missing custom document properties. | Show how to load two .xlsx files with Aspose.Cells, compare their CustomDocumentProperties collections, and output a summary of differences. | Add comprehensive error handling to a custom property comparison routine, covering file‑not‑found checks and exception reporting.
// Common Searches: aspnet compare custom document properties of two Excel files with Aspose.Cells | c# code to audit custom properties differences between two workbooks | how to list missing custom properties when comparing two .xlsx files using Aspose.Cells | generate a diff report of custom document properties in Excel workbooks with Aspose.Cells .NET
// Tags: customdocumentproperties comparison Aspose.Cells | audit workbook metadata .NET | detect mismatched custom properties Excel | compare workbook custom properties C# | list missing custom document properties Aspose

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example verifies that two .xlsx files exist, loads them into Aspose.Cells Workbook objects, extracts their CustomDocumentProperties, and iterates through each collection to identify value mismatches or properties present in only one workbook. It returns a descriptive list of differences and includes robust error handling for missing files and comparison exceptions.
class Program
{
    static void Main()
    {
        // Paths to the workbooks to compare
        string workbookPath1 = "Workbook1.xlsx";
        string workbookPath2 = "Workbook2.xlsx";

        try
        {
            // Verify that the input files exist
            if (!File.Exists(workbookPath1))
                throw new FileNotFoundException($"File not found: {workbookPath1}");
            if (!File.Exists(workbookPath2))
                throw new FileNotFoundException($"File not found: {workbookPath2}");

            // Load the two workbooks
            Workbook wb1 = new Workbook(workbookPath1);
            Workbook wb2 = new Workbook(workbookPath2);

            // Compare custom properties and collect differences
            List<string> differences = CompareCustomProperties(wb1, wb2);

            // Output the differences for auditing
            foreach (string diff in differences)
            {
                Console.WriteLine(diff);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    /// <param name="wb1">First workbook.</param>
    /// <param name="wb2">Second workbook.</param>
    /// <returns>List of difference descriptions.</returns>
    static List<string> CompareCustomProperties(Workbook wb1, Workbook wb2)
    {
        var diffs = new List<string>();

        try
        {
            // Retrieve the custom property collections
            var props1 = wb1.CustomDocumentProperties;
            var props2 = wb2.CustomDocumentProperties;

            // Check each property in the first collection
            foreach (var propObj1 in props1)
            {
                // Use dynamic to avoid compile‑time dependency on CustomDocumentProperty type
                dynamic prop1 = propObj1;
                string name = prop1.Name as string;

                if (props2.Contains(name))
                {
                    // Property exists in both workbooks; compare values
                    dynamic prop2 = props2[name];
                    if (!object.Equals(prop1.Value, prop2.Value))
                    {
                        diffs.Add($"Value mismatch for property '{name}': Workbook1 = '{prop1.Value}' vs Workbook2 = '{prop2.Value}'.");
                    }
                }
                else
                {
                    // Property missing in the second workbook
                    diffs.Add($"Property '{name}' exists in Workbook1 but not in Workbook2.");
                }
            }

            // Find properties that exist only in the second workbook
            foreach (var propObj2 in props2)
            {
                dynamic prop2 = propObj2;
                string name = prop2.Name as string;

                if (!props1.Contains(name))
                {
                    diffs.Add($"Property '{name}' exists in Workbook2 but not in Workbook1.");
                }
            }

            // If no differences were found, indicate that
            if (diffs.Count == 0)
            {
                diffs.Add("No differences in custom properties between the two workbooks.");
            }
        }
        catch (Exception ex)
        {
            diffs.Add($"Error during comparison: {ex.Message}");
        }

        return diffs;
    }
}
