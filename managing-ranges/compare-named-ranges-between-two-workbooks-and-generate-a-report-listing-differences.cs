using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

class NamedRangeComparer
{
    static void Main()
    {
        try
        {
            // Verify that the source workbooks exist
            const string file1 = "Workbook1.xlsx";
            const string file2 = "Workbook2.xlsx";

            if (!File.Exists(file1))
                throw new FileNotFoundException($"File not found: {file1}");
            if (!File.Exists(file2))
                throw new FileNotFoundException($"File not found: {file2}");

            // Load the two workbooks to be compared
            Workbook wb1 = new Workbook(file1);
            Workbook wb2 = new Workbook(file2);

            // Retrieve the collections of defined names from each workbook
            NameCollection names1 = wb1.Worksheets.Names;
            NameCollection names2 = wb2.Worksheets.Names;

            // Build dictionaries for fast lookup by name text (case‑insensitive)
            var dict1 = new Dictionary<string, Name>(StringComparer.OrdinalIgnoreCase);
            foreach (Name n in names1)
                dict1[n.Text] = n;

            var dict2 = new Dictionary<string, Name>(StringComparer.OrdinalIgnoreCase);
            foreach (Name n in names2)
                dict2[n.Text] = n;

            // Prepare a list that will hold the report lines
            List<string> reportLines = new List<string>
            {
                "Named Range Comparison Report",
                $"Generated on {DateTime.Now}",
                string.Empty
            };

            // Compare names that exist in both workbooks
            foreach (var kvp in dict1)
            {
                string name = kvp.Key;
                Name name1 = kvp.Value;

                if (dict2.TryGetValue(name, out Name name2))
                {
                    // Both workbooks contain the same named range – compare its definition
                    string ref1 = name1.RefersTo ?? string.Empty;
                    string ref2 = name2.RefersTo ?? string.Empty;

                    if (!ref1.Equals(ref2, StringComparison.OrdinalIgnoreCase))
                    {
                        reportLines.Add($"Difference in '{name}':");
                        reportLines.Add($"  Workbook1 RefersTo = {ref1}");
                        reportLines.Add($"  Workbook2 RefersTo = {ref2}");
                    }
                    else
                    {
                        // If the RefersTo strings are identical, optionally compare the actual range address
                        Aspose.Cells.Range r1 = name1.GetRange();
                        Aspose.Cells.Range r2 = name2.GetRange();

                        if (r1 != null && r2 != null && !r1.Address.Equals(r2.Address, StringComparison.OrdinalIgnoreCase))
                        {
                            reportLines.Add($"Address mismatch for '{name}':");
                            reportLines.Add($"  Workbook1 Address = {r1.Address}");
                            reportLines.Add($"  Workbook2 Address = {r2.Address}");
                        }
                    }
                }
                else
                {
                    // Named range exists only in the first workbook
                    reportLines.Add($"Only in Workbook1: {name} -> {name1.RefersTo}");
                }
            }

            // Find names that exist only in the second workbook
            foreach (var kvp in dict2)
            {
                if (!dict1.ContainsKey(kvp.Key))
                    reportLines.Add($"Only in Workbook2: {kvp.Key} -> {kvp.Value.RefersTo}");
            }

            // Create a new workbook that will hold the comparison report
            Workbook reportWb = new Workbook();
            Worksheet sheet = reportWb.Worksheets[0];

            // Write each line of the report into column A
            for (int i = 0; i < reportLines.Count; i++)
                sheet.Cells[i, 0].PutValue(reportLines[i]);

            // Save the report workbook
            const string reportFile = "NamedRangeComparisonReport.xlsx";
            reportWb.Save(reportFile);
            Console.WriteLine($"Report saved to {reportFile}");
        }
        catch (FileNotFoundException ex)
        {
            Console.Error.WriteLine($"File error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}