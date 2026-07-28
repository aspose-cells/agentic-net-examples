// Title: C# – Generate a Summary Sheet of All Named Ranges with Scope and Address using Aspose.Cells
// Description: Creates a workbook, defines workbook‑ and worksheet‑scoped named ranges, adds a "Summary" sheet, and writes each name, its scope (Workbook or specific worksheet) and the referenced address. Saves the result as NamedRangesSummary.xlsx. Demonstrates Name.GetRange, SheetIndex and error‑fallback handling in Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# named ranges | list named ranges Aspose.Cells | named range scope .NET | GetRange address Aspose | summary worksheet Excel | Workbook‑scoped vs Worksheet‑scoped names | C# Excel automation Aspose | export named ranges to Excel
// Common Searches: how to list all named ranges in Aspose.Cells | C# code to show named range scope and address | Aspose.Cells create summary sheet of names | retrieve address of a named range with Aspose.Cells | differentiate workbook and worksheet scoped names in .NET
// Developer Intent: Add a worksheet that enumerates every defined name, indicates whether it is workbook‑scoped or tied to a specific sheet, and displays the range address.
// Use Cases: Provide end‑users with a reference sheet documenting all named ranges in a workbook. | Validate that required named ranges exist with the correct scope before running data‑processing logic. | Generate an audit report of named ranges for compliance, documentation, or debugging purposes.
// AI Prompts: Write C# code with Aspose.Cells that adds a "Summary" sheet listing each defined name, its scope (workbook or worksheet), and its address. | Explain how the SheetIndex property distinguishes workbook‑scoped from worksheet‑scoped names in Aspose.Cells. | Suggest robust error‑handling patterns when calling Name.GetRange() for names that may reference formulas or invalid ranges.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Creates a workbook, defines workbook‑ and worksheet‑scoped named ranges, adds a "Summary" sheet, and writes each name, its scope (Workbook or specific worksheet) and the referenced address. Saves the result as NamedRangesSummary.xlsx. Demonstrates Name.GetRange, SheetIndex and error‑fallback handling in Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // Sample data: create some named ranges for demo
            // -------------------------------------------------
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";
            sheet1.Cells["A1"].PutValue(1);
            sheet1.Cells["A2"].PutValue(2);
            sheet1.Cells["A3"].PutValue(3);

            // Global (workbook‑scoped) named range
            int globalIdx = workbook.Worksheets.Names.Add("GlobalRange");
            workbook.Worksheets.Names[globalIdx].RefersTo = "=Sheet1!$A$1:$A$3";

            // Worksheet‑scoped named range on Sheet2
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
            sheet2.Cells["B1"].PutValue(10);
            sheet2.Cells["B2"].PutValue(20);
            int localIdx = workbook.Worksheets.Names.Add("Sheet2!LocalRange");
            Name localName = workbook.Worksheets.Names[localIdx];
            localName.RefersTo = "=Sheet2!$B$1:$B$2";
            // Set the scope to the second worksheet (one‑based index)
            localName.SheetIndex = sheet2.Index + 1;

            // -------------------------------------------------
            // Create the summary worksheet
            // -------------------------------------------------
            Worksheet summary = workbook.Worksheets.Add("Summary");
            // Header row
            summary.Cells["A1"].PutValue("Name");
            summary.Cells["B1"].PutValue("Scope");
            summary.Cells["C1"].PutValue("Address");

            int currentRow = 1; // zero‑based index; row 1 is the second row in the sheet

            // -------------------------------------------------
            // Iterate over all defined names and write details
            // -------------------------------------------------
            foreach (Name name in workbook.Worksheets.Names)
            {
                // Determine the scope of the name
                string scope = name.SheetIndex == 0
                    ? "Workbook"
                    : $"Worksheet ({workbook.Worksheets[name.SheetIndex - 1].Name})";

                // Retrieve the address of the range the name refers to
                string address;
                try
                {
                    AsposeRange range = name.GetRange();
                    address = range != null ? range.Address : name.RefersTo;
                }
                catch
                {
                    // Fallback to the RefersTo formula if GetRange fails
                    address = name.RefersTo;
                }

                // Write the information into the summary sheet
                summary.Cells[currentRow, 0].PutValue(name.Text);   // Name
                summary.Cells[currentRow, 1].PutValue(scope);      // Scope
                summary.Cells[currentRow, 2].PutValue(address);    // Address

                currentRow++;
            }

            // -------------------------------------------------
            // Save the workbook with the summary sheet
            // -------------------------------------------------
            workbook.Save("NamedRangesSummary.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
