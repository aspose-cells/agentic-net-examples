// Title: C# – Generate a Summary Sheet of All Named Ranges with Addresses and RefersTo Formulas using Aspose.Cells
// Description: This example creates a workbook, defines sample named ranges, adds a "Summary" worksheet, and writes a table that shows each defined name, the range address (or N/A), and the original RefersTo formula. The workbook is then saved as NamedRangeSummary.xlsx.
// Keywords: Aspose.Cells C# | list named ranges | named range address | RefersTo formula | GetRange Aspose.Cells | Excel named range summary | iterate workbook names | export named ranges | Aspose.Cells API | C# Excel automation
// Common Searches: how to list all named ranges with Aspose.Cells | C# get address of a named range in Excel | Aspose.Cells create summary sheet of names | retrieve RefersTo formula using Aspose.Cells | iterate over defined names in a workbook C#
// Developer Intent: Create a worksheet that enumerates every defined name, displaying its name, cell address (when available), and the RefersTo expression.
// Use Cases: Document and audit all named ranges in a generated report. | Provide end‑users a quick reference of workbook names for troubleshooting. | Validate that named ranges point to the correct cells before distribution.
// AI Prompts: Write C# code with Aspose.Cells that builds a summary sheet of all workbook names, showing each name, its address, and its RefersTo formula, handling non‑range names gracefully. | Explain how to obtain a Range object from a Name and read its Address property in Aspose.Cells. | Show how to export the named‑range summary to a CSV file instead of an Excel worksheet using Aspose.Cells.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsNamedRangeSummary
{
    // This example creates a workbook, defines sample named ranges, adds a "Summary" worksheet, and writes a table that shows each defined name, the range address (or N/A), and the original RefersTo formula. The workbook is then saved as NamedRangeSummary.xlsx.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle create)
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // Sample data: create some named ranges for demo
                // -------------------------------------------------
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "Data";

                // Populate some cells
                sheet1.Cells["A1"].PutValue("Item");
                sheet1.Cells["B1"].PutValue("Quantity");
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

                // Iterate over all defined names
                int row = 1; // start after header
                foreach (Name name in workbook.Worksheets.Names)
                {
                    // Name text
                    summary.Cells[row, 0].PutValue(name.Text);

                    // Attempt to get the range the name refers to
                    AsposeRange range = null;
                    try
                    {
                        range = name.GetRange(); // may return null if not a range
                    }
                    catch
                    {
                        // ignore exceptions; keep range null
                    }

                    // Address (if range is available)
                    string address = range != null ? range.Address : "N/A";
                    summary.Cells[row, 1].PutValue(address);

                    // RefersTo formula (as stored in the Name object)
                    summary.Cells[row, 2].PutValue(name.RefersTo);

                    row++;
                }

                // Save the workbook (lifecycle save)
                workbook.Save("NamedRangeSummary.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
