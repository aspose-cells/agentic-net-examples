// Title: Add a left‑aligned date footer to every worksheet in an Excel workbook with Aspose.Cells for .NET
// Description: Loads an existing workbook, loops through all worksheets, and uses PageSetup.SetFooter(0, "&D") to place the current date in the left footer of each sheet before saving the file.
// Keywords: Aspose.Cells left footer date | C# add footer to all worksheets | PageSetup SetFooter example | Excel left‑aligned footer .NET | Aspose.Cells insert current date footer
// Common Searches: Aspose.Cells add left footer date to each sheet | C# set footer for all worksheets in Excel | How to use PageSetup.SetFooter with &D | Add date footer to workbook using Aspose.Cells
// Developer Intent: Insert the current date as a left‑aligned footer on every worksheet of an existing Excel file.
// Use Cases: Automatically timestamp printed reports across all sheets. | Ensure consistent footers for multi‑sheet financial statements. | Add issuance dates to invoices generated in bulk.
// AI Prompts: Generate C# code to add a right‑aligned page‑number footer to all worksheets with Aspose.Cells. | Show how to combine workbook name and current date in a centered footer using SetFooter. | Explain the three footer sections (left, center, right) and formatting options in Aspose.Cells.

using System;
using Aspose.Cells;

// Loads an existing workbook, loops through all worksheets, and uses PageSetup.SetFooter(0, "&D") to place the current date in the left footer of each sheet before saving the file.
class AddFooterExample
{
    static void Main()
    {
        // Load the existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Add a left-aligned footer with the current date to each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Section 0 = left footer, &D inserts the current date
            sheet.PageSetup.SetFooter(0, "&D");
        }

        // Save the modified workbook (replace with desired output path)
        workbook.Save("output.xlsx");
    }
}
