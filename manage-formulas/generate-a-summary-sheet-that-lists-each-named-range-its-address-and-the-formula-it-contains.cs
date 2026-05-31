using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class NamedRangeSummary
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle: create)
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // Sample data: create some named ranges for demo
                // -------------------------------------------------
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "Data";

                // Fill some cells
                sheet1.Cells["A1"].PutValue("Item");
                sheet1.Cells["B1"].PutValue("Qty");
                sheet1.Cells["A2"].PutValue("Apple");
                sheet1.Cells["B2"].PutValue(10);
                sheet1.Cells["A3"].PutValue("Banana");
                sheet1.Cells["B3"].PutValue(20);

                // Define named ranges
                int idx1 = workbook.Worksheets.Names.Add("Items");
                workbook.Worksheets.Names[idx1].RefersTo = "=Data!$A$2:$A$3";

                int idx2 = workbook.Worksheets.Names.Add("Quantities");
                workbook.Worksheets.Names[idx2].RefersTo = "=Data!$B$2:$B$3";

                // -------------------------------------------------
                // Create a summary worksheet
                // -------------------------------------------------
                Worksheet summary = workbook.Worksheets.Add("Summary");

                // Write header row
                summary.Cells[0, 0].PutValue("Name");
                summary.Cells[0, 1].PutValue("Address");
                summary.Cells[0, 2].PutValue("RefersTo Formula");

                // Iterate over all defined names in the workbook
                int row = 1; // start after header
                foreach (Name name in workbook.Worksheets.Names)
                {
                    // Write the name text
                    summary.Cells[row, 0].PutValue(name.Text);

                    // Try to obtain the actual range (if the name refers to a range)
                    Aspose.Cells.Range range = null;
                    try
                    {
                        range = name.GetRange(); // may return null if not a range
                    }
                    catch
                    {
                        // ignore exceptions; keep range null
                    }

                    // Write address if available
                    summary.Cells[row, 1].PutValue(range != null ? range.Address : "N/A");

                    // Write the RefersTo formula (the definition of the name)
                    summary.Cells[row, 2].PutValue(name.RefersTo);

                    row++;
                }

                // -------------------------------------------------
                // Save the workbook (lifecycle: save)
                // -------------------------------------------------
                workbook.Save("NamedRangeSummary.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Runtime error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            NamedRangeSummary.Run();
        }
    }
}