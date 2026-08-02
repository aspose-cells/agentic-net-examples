using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Properties; // For PropertyType

namespace AsposeCellsCustomPropertyComparison
{
    class Program
    {
        static void Main(string[] args)
        {
            // Paths to the two workbooks to compare
            string workbookPath1 = "Workbook1.xlsx";
            string workbookPath2 = "Workbook2.xlsx";

            try
            {
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

                // Retrieve custom document properties collections
                var props1 = wb1.CustomDocumentProperties;
                var props2 = wb2.CustomDocumentProperties;

                // Build dictionaries for fast lookup by property name (case‑insensitive)
                var dict1 = new Dictionary<string, (object Value, PropertyType Type)>(StringComparer.OrdinalIgnoreCase);
                foreach (DocumentProperty prop in props1)
                {
                    dict1[prop.Name] = (prop.Value, prop.Type);
                }

                var dict2 = new Dictionary<string, (object Value, PropertyType Type)>(StringComparer.OrdinalIgnoreCase);
                foreach (DocumentProperty prop in props2)
                {
                    dict2[prop.Name] = (prop.Value, prop.Type);
                }

                // List to hold difference descriptions
                var differences = new List<string>();

                // Check properties present in the first workbook
                foreach (var kvp in dict1)
                {
                    string name = kvp.Key;
                    var (value1, type1) = kvp.Value;

                    if (!dict2.TryGetValue(name, out var valueType2))
                    {
                        differences.Add($"Only in first workbook: '{name}' = {value1} (Type: {type1})");
                    }
                    else
                    {
                        var (value2, type2) = valueType2;
                        bool valuesEqual = Equals(value1, value2) && type1 == type2;
                        if (!valuesEqual)
                        {
                            differences.Add($"Changed property '{name}': First = {value1} (Type: {type1}), Second = {value2} (Type: {type2})");
                        }
                    }
                }

                // Check properties present only in the second workbook
                foreach (var kvp in dict2)
                {
                    if (!dict1.ContainsKey(kvp.Key))
                    {
                        var (value2, type2) = kvp.Value;
                        differences.Add($"Only in second workbook: '{kvp.Key}' = {value2} (Type: {type2})");
                    }
                }

                // Output the differences
                Console.WriteLine("Custom Document Property Differences:");
                if (differences.Count == 0)
                {
                    Console.WriteLine("No differences found.");
                }
                else
                {
                    foreach (string diff in differences)
                        Console.WriteLine(diff);
                }

                // Save the differences to a text file for auditing
                try
                {
                    string reportPath = "CustomPropertyDifferences.txt";
                    File.WriteAllLines(reportPath, differences);
                    Console.WriteLine($"Difference report saved to: {reportPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to write report file: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}