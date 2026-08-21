// Title: C# – Compare Named Ranges in Two Excel Workbooks and Export a Difference Report with Aspose.Cells
// Description: Loads two workbooks, extracts their NameCollection objects, builds a case‑insensitive lookup, and identifies missing named ranges or RefersTo mismatches. The findings are written to a new workbook with headers, auto‑fitted columns, and saved as a comparison report.
// Keywords: Aspose.Cells compare named ranges | C# named range difference report | Excel workbook defined names comparison | Aspose.Cells NameCollection .NET | detect missing named ranges | RefersTo formula mismatch | generate Excel audit report | compare two workbooks programmatically
// Common Searches: compare named ranges Aspose.Cells C# | generate report of Excel defined name differences | find missing named ranges between two workbooks | Aspose.Cells detect RefersTo changes | C# code to audit named ranges in Excel files
// Developer Intent: Find and document discrepancies in named ranges across two Excel files using Aspose.Cells.
// Use Cases: Track changes to named range references between versioned workbooks. | Identify ranges that exist only in one of the files for data migration. | Create an audit log of named‑range modifications for compliance reviews.
// AI Prompts: Generate C# code with Aspose.Cells that compares named ranges in two workbooks and outputs a CSV summary. | Explain how to add the worksheet scope of each named range to the comparison report. | Suggest enhancements to also compare comments or descriptions attached to named ranges.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace NamedRangeComparison
{
    // Loads two workbooks, extracts their NameCollection objects, builds a case‑insensitive lookup, and identifies missing named ranges or RefersTo mismatches. The findings are written to a new workbook with headers, auto‑fitted columns, and saved as a comparison report.
    class Program
    {
        static void Main(string[] args)
        {
            // Load the two workbooks to compare
            Workbook wb1 = new Workbook("Workbook1.xlsx");
            Workbook wb2 = new Workbook("Workbook2.xlsx");

            // Get the collections of defined names from each workbook
            NameCollection names1 = wb1.Worksheets.Names;
            NameCollection names2 = wb2.Worksheets.Names;

            // Build a lookup dictionary for the names in the second workbook
            var dictNames2 = new Dictionary<string, Name>(StringComparer.OrdinalIgnoreCase);
            foreach (Name n in names2)
            {
                dictNames2[n.Text] = n;
            }

            // Prepare a list to hold report rows
            var reportRows = new List<string[]>();

            // Compare names present in the first workbook
            foreach (Name n1 in names1)
            {
                string nameText = n1.Text;
                if (dictNames2.TryGetValue(nameText, out Name n2))
                {
                    // Name exists in both workbooks; compare the RefersTo formulas
                    string ref1 = n1.RefersTo ?? string.Empty;
                    string ref2 = n2.RefersTo ?? string.Empty;
                    if (!string.Equals(ref1, ref2, StringComparison.OrdinalIgnoreCase))
                    {
                        reportRows.Add(new[] { nameText, ref1, ref2, "RefersTo differs" });
                    }
                    // Remove the matched entry so we can later detect names only in wb2
                    dictNames2.Remove(nameText);
                }
                else
                {
                    // Name missing in the second workbook
                    reportRows.Add(new[] { nameText, n1.RefersTo ?? string.Empty, "", "Missing in Workbook2" });
                }
            }

            // Any remaining names in dictNames2 are missing from the first workbook
            foreach (var kvp in dictNames2)
            {
                Name n2 = kvp.Value;
                reportRows.Add(new[] { n2.Text, "", n2.RefersTo ?? string.Empty, "Missing in Workbook1" });
            }

            // Create a new workbook to hold the comparison report
            Workbook reportWb = new Workbook();
            Worksheet sheet = reportWb.Worksheets[0];

            // Write header row
            sheet.Cells[0, 0].PutValue("Name");
            sheet.Cells[0, 1].PutValue("Workbook1 RefersTo");
            sheet.Cells[0, 2].PutValue("Workbook2 RefersTo");
            sheet.Cells[0, 3].PutValue("Status");

            // Populate the report rows
            for (int i = 0; i < reportRows.Count; i++)
            {
                string[] row = reportRows[i];
                sheet.Cells[i + 1, 0].PutValue(row[0]);
                sheet.Cells[i + 1, 1].PutValue(row[1]);
                sheet.Cells[i + 1, 2].PutValue(row[2]);
                sheet.Cells[i + 1, 3].PutValue(row[3]);
            }

            // Auto-fit columns for better readability
            sheet.AutoFitColumns();

            // Save the report workbook
            reportWb.Save("NamedRangeComparisonReport.xlsx");
        }
    }
}
