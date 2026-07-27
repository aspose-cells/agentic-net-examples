using System;
using Aspose.Cells;

namespace NamedRangeComparison
{
    class Program
    {
        static void Main()
        {
            // Load the two workbooks to compare
            Workbook wb1 = new Workbook("Workbook1.xlsx");   // first workbook
            Workbook wb2 = new Workbook("Workbook2.xlsx");   // second workbook

            // Get the collections of defined names from each workbook
            NameCollection names1 = wb1.Worksheets.Names;
            NameCollection names2 = wb2.Worksheets.Names;

            // Create a new workbook that will hold the comparison report
            Workbook reportWb = new Workbook();
            Worksheet reportSheet = reportWb.Worksheets[0];
            reportSheet.Name = "NamedRangeDifferences";

            // Write header row
            reportSheet.Cells["A1"].PutValue("Name");
            reportSheet.Cells["B1"].PutValue("Workbook1 RefersTo");
            reportSheet.Cells["C1"].PutValue("Workbook2 RefersTo");
            reportSheet.Cells["D1"].PutValue("Status");

            int reportRow = 1; // zero‑based index; row 1 is the second row (after header)

            // Helper to add a row to the report
            void AddReportRow(string name, string ref1, string ref2, string status)
            {
                reportSheet.Cells[reportRow, 0].PutValue(name);
                reportSheet.Cells[reportRow, 1].PutValue(ref1);
                reportSheet.Cells[reportRow, 2].PutValue(ref2);
                reportSheet.Cells[reportRow, 3].PutValue(status);
                reportRow++;
            }

            // Process names that exist in the first workbook
            foreach (Name name1 in names1)
            {
                string nameText = name1.Text;
                string ref1 = name1.RefersTo ?? string.Empty;

                // Try to find the same name in the second workbook
                Name name2 = names2[nameText];
                if (name2 != null)
                {
                    string ref2 = name2.RefersTo ?? string.Empty;
                    string status = string.Equals(ref1, ref2, StringComparison.OrdinalIgnoreCase)
                                    ? "Same"
                                    : "Different";
                    AddReportRow(nameText, ref1, ref2, status);
                }
                else
                {
                    // Name missing in workbook2
                    AddReportRow(nameText, ref1, "", "Missing in Workbook2");
                }
            }

            // Process names that exist only in the second workbook
            foreach (Name name2 in names2)
            {
                string nameText = name2.Text;
                // If already reported (exists in workbook1) skip
                if (names1[nameText] != null) continue;

                string ref2 = name2.RefersTo ?? string.Empty;
                AddReportRow(nameText, "", ref2, "Missing in Workbook1");
            }

            // Auto‑fit columns for better readability
            reportSheet.AutoFitColumns();

            // Save the report workbook
            reportWb.Save("NamedRangeComparisonReport.xlsx");
        }
    }
}