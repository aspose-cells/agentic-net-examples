// Title: C# code to compare custom document properties of two Excel workbooks and generate an audit report with Aspose.Cells
// Common Searches: C# compare custom document properties of two Excel files using Aspose.Cells | how to create a custom property differences report for workbooks in .NET | Aspose.Cells detect added modified removed custom properties between Excel workbooks | generate audit worksheet for Excel workbook metadata differences with C# | compare workbook custom properties and export differences to a new Excel file
// Tags: custom document property comparison Aspose.Cells C# | Excel workbook metadata audit report .NET | detect added modified removed properties Aspose.Cells | write property differences to worksheet C# | case-insensitive property dictionary Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads two Excel workbooks, extracts their CustomDocumentProperties, builds case‑insensitive dictionaries, identifies added, modified, and removed properties, writes each difference as a text line into a new worksheet named "Custom Property Differences", and saves the audit report as an Xlsx file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Paths to the workbooks to compare
            string workbookPath1 = "Workbook1.xlsx";
            string workbookPath2 = "Workbook2.xlsx";

            // Verify that source files exist
            if (!File.Exists(workbookPath1))
                throw new FileNotFoundException($"File not found: {workbookPath1}");
            if (!File.Exists(workbookPath2))
                throw new FileNotFoundException($"File not found: {workbookPath2}");

            // Load the workbooks
            Workbook wb1 = new Workbook(workbookPath1);
            Workbook wb2 = new Workbook(workbookPath2);

            // Retrieve custom document properties from each workbook
            var props1 = wb1.CustomDocumentProperties;
            var props2 = wb2.CustomDocumentProperties;

            // Build dictionaries for quick lookup (case‑insensitive keys)
            var dict1 = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            foreach (var prop in props1)
            {
                dict1[prop.Name] = prop.Value;
            }

            var dict2 = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            foreach (var prop in props2)
            {
                dict2[prop.Name] = prop.Value;
            }

            // Collect differences
            List<string> differences = new List<string>();

            // Detect added or modified properties
            foreach (var kvp in dict2)
            {
                if (!dict1.ContainsKey(kvp.Key))
                {
                    differences.Add($"Added: {kvp.Key} = {kvp.Value}");
                }
                else
                {
                    object oldValue = dict1[kvp.Key];
                    if (!object.Equals(oldValue, kvp.Value))
                    {
                        differences.Add($"Modified: {kvp.Key} from '{oldValue}' to '{kvp.Value}'");
                    }
                }
            }

            // Detect removed properties
            foreach (var kvp in dict1)
            {
                if (!dict2.ContainsKey(kvp.Key))
                {
                    differences.Add($"Removed: {kvp.Key} = {kvp.Value}");
                }
            }

            // Create a new workbook to hold the audit report
            Workbook reportWb = new Workbook();
            Worksheet sheet = reportWb.Worksheets[0];
            sheet.Name = "Custom Property Differences";

            // Write header
            sheet.Cells[0, 0].PutValue("Difference");

            // Write each difference line
            for (int i = 0; i < differences.Count; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(differences[i]);
            }

            // Save the audit report
            string reportPath = "CustomPropertiesDifferences.xlsx";
            reportWb.Save(reportPath, SaveFormat.Xlsx);
            Console.WriteLine($"Report saved to {reportPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
