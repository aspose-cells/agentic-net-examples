// Title: C# – Export All Named Ranges and Their RefersTo Formulas to a Summary Sheet with Aspose.Cells
// Description: Creates a workbook, defines sample named ranges, adds a "Summary" worksheet, writes column headers, retrieves and optionally sorts the NameCollection, then records each named range together with its RefersTo formula before saving the file as NamedRangesSummary.xlsx.
// Keywords: Aspose.Cells | C# export named ranges | list defined names | RefersTo formula | Excel summary worksheet | NameCollection sort | named range audit | Excel workbook structure
// Common Searches: Aspose.Cells list named ranges C# | Export named ranges to summary sheet Aspose | Get RefersTo formula of named range using Aspose.Cells | Sort defined names before exporting Aspose | Create summary worksheet with named ranges Aspose.Cells
// Developer Intent: Create a worksheet that lists every named range in the workbook along with its RefersTo expression.
// Use Cases: Audit workbook definitions by exporting all named ranges and formulas. | Provide end‑users a reference sheet showing how ranges are defined. | Generate a sorted catalog of named ranges for documentation or debugging. | Automate reporting of dynamic range definitions in generated Excel files.
// AI Prompts: Generate C# code with Aspose.Cells that iterates through workbook.Worksheets.Names, extracts each Name.Text and Name.RefersTo, writes them to a new sheet called 'Summary' with headers, sorts the collection alphabetically, and saves the file. | Show how to add a header row and format the columns when exporting named ranges and their formulas using Aspose.Cells for .NET. | Explain how to retrieve and optionally sort the NameCollection before writing to a summary worksheet in Aspose.Cells.

using System;
using Aspose.Cells;

namespace NamedRangeExportExample
{
    // Creates a workbook, defines sample named ranges, adds a "Summary" worksheet, writes column headers, retrieves and optionally sorts the NameCollection, then records each named range together with its RefersTo formula before saving the file as NamedRangesSummary.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // OPTIONAL: create some sample named ranges
            // -------------------------------------------------
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Data";

            // Fill sample data
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
            // Create a summary worksheet to list named ranges
            // -------------------------------------------------
            Worksheet summary = workbook.Worksheets.Add("Summary");

            // Write header
            summary.Cells["A1"].PutValue("Named Range");
            summary.Cells["B1"].PutValue("Refers To Formula");

            // Retrieve all defined names
            NameCollection names = workbook.Worksheets.Names;

            // Optional: sort names for a tidy output
            names.Sort();

            // Export each name and its formula to the summary sheet
            int row = 1; // zero‑based index; row 1 is the second row (after header)
            foreach (Name name in names)
            {
                // Write the name text
                summary.Cells[row, 0].PutValue(name.Text);
                // Write the formula the name refers to (including leading '=')
                summary.Cells[row, 1].PutValue(name.RefersTo);
                row++;
            }

            // Save the workbook (lifecycle: save)
            workbook.Save("NamedRangesSummary.xlsx");
        }
    }
}
