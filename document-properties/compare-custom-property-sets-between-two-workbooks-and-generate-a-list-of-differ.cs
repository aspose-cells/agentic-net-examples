using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsCustomPropertyComparison
{
    class Program
    {
        static void Main(string[] args)
        {
            // Paths to the two workbooks to compare
            string workbookPath1 = "Workbook1.xlsx";
            string workbookPath2 = "Workbook2.xlsx";

            // Load the workbooks (load rule)
            Workbook wb1 = new Workbook(workbookPath1);
            Workbook wb2 = new Workbook(workbookPath2);

            // Retrieve custom document properties collections
            CustomDocumentPropertyCollection props1 = wb1.CustomDocumentProperties;
            CustomDocumentPropertyCollection props2 = wb2.CustomDocumentProperties;

            // Build dictionaries for quick lookup (key = property name)
            var dict1 = new Dictionary<string, DocumentProperty>(StringComparer.OrdinalIgnoreCase);
            foreach (DocumentProperty prop in props1)
                dict1[prop.Name] = prop;

            var dict2 = new Dictionary<string, DocumentProperty>(StringComparer.OrdinalIgnoreCase);
            foreach (DocumentProperty prop in props2)
                dict2[prop.Name] = prop;

            // List to hold difference descriptions
            List<string> differences = new List<string>();

            // Compare properties present in workbook 1
            foreach (var kvp in dict1)
            {
                string name = kvp.Key;
                DocumentProperty prop1 = kvp.Value;

                if (!dict2.TryGetValue(name, out DocumentProperty prop2))
                {
                    // Property exists only in workbook 1
                    differences.Add($"Only in Workbook1: {name} = {prop1.Value} ({prop1.Type})");
                }
                else
                {
                    // Property exists in both – compare value and type
                    bool valuesEqual = Equals(prop1.Value, prop2.Value);
                    bool typesEqual = prop1.Type == prop2.Type;

                    if (!valuesEqual || !typesEqual)
                    {
                        differences.Add($"Different for '{name}': " +
                                        $"Workbook1 = {prop1.Value} ({prop1.Type}), " +
                                        $"Workbook2 = {prop2.Value} ({prop2.Type})");
                    }
                }
            }

            // Find properties that exist only in workbook 2
            foreach (var kvp in dict2)
            {
                if (!dict1.ContainsKey(kvp.Key))
                {
                    DocumentProperty prop2 = kvp.Value;
                    differences.Add($"Only in Workbook2: {kvp.Key} = {prop2.Value} ({prop2.Type})");
                }
            }

            // If no differences were found
            if (differences.Count == 0)
                differences.Add("No differences in custom document properties.");

            // Create a new workbook to store the audit report (create rule)
            Workbook reportWorkbook = new Workbook();
            Worksheet sheet = reportWorkbook.Worksheets[0];

            // Write header
            sheet.Cells["A1"].PutValue("Custom Document Property Differences");
            sheet.Cells["A2"].PutValue("---------------------------------------------------");

            // Write each difference line into successive rows
            for (int i = 0; i < differences.Count; i++)
            {
                sheet.Cells[i + 3, 0].PutValue(differences[i]); // Row index i+3 (0‑based), column 0 (A)
            }

            // Save the report workbook (save rule)
            reportWorkbook.Save("CustomPropertyDifferences.xlsx");

            // Also output to console for immediate visibility
            Console.WriteLine("Audit report generated: CustomPropertyDifferences.xlsx");
            foreach (string line in differences)
                Console.WriteLine(line);
        }
    }
}