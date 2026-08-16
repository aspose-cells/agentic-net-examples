// Title: Generate a Summary Sheet of All Named Ranges with Scope and Address using Aspose.Cells for .NET
// Description: C# example that creates a workbook, adds worksheet‑scoped and workbook‑scoped named ranges, then builds a "Summary" worksheet. The code writes headers (Name, Scope, Address), iterates through the workbook's NameCollection, determines each name's scope, extracts the range address via GetRange or RefersTo, and saves the result as NamedRangesSummary.xlsx.
// Keywords: Aspose.Cells C# named ranges | list named ranges .NET | named range scope Aspose | Excel summary sheet programmatically | retrieve range address Aspose.Cells | global named range workbook | worksheet scoped named range | NameCollection Aspose | export defined names to sheet | C# Excel automation
// Common Searches: how to list all named ranges in an Excel file using Aspose.Cells | C# code to get named range scope and address | create a report of defined names with Aspose.Cells for .NET | Aspose.Cells retrieve workbook scoped names | generate summary worksheet of named ranges
// Developer Intent: Provide a ready‑to‑run C# snippet that enumerates every defined name in a workbook, identifies whether it is workbook‑scoped or sheet‑scoped, captures its address, and writes this information to a new summary sheet.
// Use Cases: Audit and document all named ranges before delivering a workbook to stakeholders. | Give end‑users a quick reference sheet showing where each defined name points. | Validate that global and sheet‑level names reference the correct cells during automated testing.
// AI Prompts: Write C# code to add a new worksheet‑scoped named range and automatically update the existing summary sheet. | Extend the example to include a fourth column that displays the full RefersTo formula for each named range. | Create unit tests that verify the summary worksheet contains the correct name, scope, and address for every defined name.

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    // C# example that creates a workbook, adds worksheet‑scoped and workbook‑scoped named ranges, then builds a "Summary" worksheet. The code writes headers (Name, Scope, Address), iterates through the workbook's NameCollection, determines each name's scope, extracts the range address via GetRange or RefersTo, and saves the result as NamedRangesSummary.xlsx.
    class SummaryNamedRanges
    {
        static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample data and named ranges on the first worksheet
            Worksheet ws1 = workbook.Worksheets[0];
            ws1.Name = "Sheet1";
            ws1.Cells["A1"].PutValue(1);
            ws1.Cells["A2"].PutValue(2);
            ws1.Cells["B1"].PutValue(3);
            ws1.Cells["B2"].PutValue(4);
            ws1.Cells.CreateRange("A1:B2").Name = "RangeSheet1";

            // Add a second worksheet with its own named range
            Worksheet ws2 = workbook.Worksheets.Add("Sheet2");
            ws2.Cells["C3"].PutValue(5);
            ws2.Cells["D4"].PutValue(6);
            ws2.Cells.CreateRange("C3:D4").Name = "RangeSheet2";

            // Add a workbook‑scoped (global) named range
            int globalIndex = workbook.Worksheets.Names.Add("GlobalRange");
            workbook.Worksheets.Names[globalIndex].RefersTo = "=Sheet1!$A$1";

            // Create a summary worksheet
            Worksheet summary = workbook.Worksheets.Add("Summary");
            // Write header row
            summary.Cells["A1"].PutValue("Name");
            summary.Cells["B1"].PutValue("Scope");
            summary.Cells["C1"].PutValue("Address");

            // Retrieve all defined names
            NameCollection names = workbook.Worksheets.Names;
            int row = 2; // start after header

            foreach (Name name in names)
            {
                // Column A: name text
                summary.Cells[row, 0].PutValue(name.Text);

                // Column B: scope (Workbook or specific worksheet name)
                string scope;
                if (name.SheetIndex == -1) // -1 indicates workbook scope
                    scope = "Workbook";
                else
                    scope = workbook.Worksheets[name.SheetIndex].Name; // SheetIndex is zero‑based for worksheet scope
                summary.Cells[row, 1].PutValue(scope);

                // Column C: address of the range
                string address;
                try
                {
                    AsposeRange rng = name.GetRange();
                    address = rng != null ? rng.Address : name.RefersTo;
                }
                catch
                {
                    // If GetRange fails, fall back to RefersTo string
                    address = name.RefersTo;
                }
                summary.Cells[row, 2].PutValue(address);

                row++;
            }

            // Save the workbook with the summary sheet
            string outputPath = "NamedRangesSummary.xlsx";
            workbook.Save(outputPath);
        }
    }
}
