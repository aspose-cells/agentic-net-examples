// Title: Enumerate Named Ranges with Addresses and RefersTo Formulas using Aspose.Cells for .NET
// Description: This C# example creates a workbook, defines sample data and two named ranges, then adds a "Summary" worksheet. It writes column headers and loops through every defined name, extracting the name text, the range address (or "N/A" when unavailable), and the RefersTo expression, finally saving the file as NamedRangeSummary.xlsx.
// Keywords: Aspose.Cells C# list named ranges | named range address .NET | RefersTo formula extraction | enumerate defined names workbook | Aspose.Cells summary sheet | C# Excel automation | global developers
// Common Searches: how to list all named ranges in Aspose.Cells | C# get address of a defined name Excel | Aspose.Cells create summary worksheet for named ranges | retrieve RefersTo formula with Aspose.Cells .NET | export named range metadata to new sheet
// Developer Intent: Produce a worksheet that catalogs each defined name, its cell address (or placeholder), and its RefersTo formula.
// Use Cases: Generate a quick‑reference guide for end users to locate and understand workbook named ranges. | Audit and validate data models by reporting all defined names and their targets. | Feed named‑range metadata into documentation pipelines or reporting tools.
// AI Prompts: Write C# code with Aspose.Cells that adds a "Summary" sheet listing every defined name, its address (or "N/A"), and its RefersTo formula, while safely handling non‑range names. | Explain the behavior of Name.GetRange() when a name points to a constant, a formula, or an external reference in Aspose.Cells. | Modify the sample to also display the total number of cells covered by each named range on the summary sheet.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    // This C# example creates a workbook, defines sample data and two named ranges, then adds a "Summary" worksheet. It writes column headers and loops through every defined name, extracting the name text, the range address (or "N/A" when unavailable), and the RefersTo expression, finally saving the file as NamedRangeSummary.xlsx.
    public class NamedRangeSummary
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle: create)
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // Sample data and named ranges (for demonstration)
                // -------------------------------------------------
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "Data";

                // Populate some cells
                sheet1.Cells["A1"].PutValue(10);
                sheet1.Cells["A2"].PutValue(20);
                sheet1.Cells["A3"].PutValue(30);
                sheet1.Cells["B1"].PutValue("Alpha");
                sheet1.Cells["B2"].PutValue("Beta");
                sheet1.Cells["B3"].PutValue("Gamma");

                // Create named ranges
                int idx1 = workbook.Worksheets.Names.Add("Numbers");
                workbook.Worksheets.Names[idx1].RefersTo = "=Data!$A$1:$A$3";

                int idx2 = workbook.Worksheets.Names.Add("Texts");
                workbook.Worksheets.Names[idx2].RefersTo = "=Data!$B$1:$B$3";

                // -------------------------------------------------
                // Create a summary sheet
                // -------------------------------------------------
                Worksheet summary = workbook.Worksheets.Add("Summary");

                // Write header row
                summary.Cells["A1"].PutValue("Name");
                summary.Cells["B1"].PutValue("Address");
                summary.Cells["C1"].PutValue("RefersTo Formula");

                // Iterate over all defined names
                int row = 2; // start after header
                foreach (Name name in workbook.Worksheets.Names)
                {
                    // Name text
                    summary.Cells[row, 0].PutValue(name.Text);

                    // Attempt to get the range the name refers to
                    AsposeRange range = null;
                    try
                    {
                        // GetRange() returns null if the name does not refer to a range
                        range = name.GetRange();
                    }
                    catch
                    {
                        // Ignored – some names may refer to formulas or external links
                    }

                    // Address (if range is available)
                    string address = range != null ? range.Address : "N/A";
                    summary.Cells[row, 1].PutValue(address);

                    // The formula (RefersTo) that defines the name
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
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            NamedRangeSummary.Run();
        }
    }
}
