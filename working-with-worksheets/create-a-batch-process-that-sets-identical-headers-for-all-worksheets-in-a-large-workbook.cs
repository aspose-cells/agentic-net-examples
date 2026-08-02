// Title: C# – Apply identical headers to all worksheets using Aspose.Cells
// Description: Load a workbook, define left, center, and right header scripts (file name, page number, date), loop through each worksheet, set the headers with PageSetup.SetHeader, and save the updated file.
// Keywords: Aspose.Cells C# header | set worksheet header batch | PageSetup SetHeader | uniform Excel headers | apply same header to all sheets
// Common Searches: Aspose.Cells set same header for every sheet | C# batch update Excel worksheet headers | How to apply identical page header to all worksheets in .NET
// Developer Intent: Programmatically assign a consistent header to every worksheet in a large Excel workbook.
// Use Cases: Add file name, page number, and date to each sheet before printing. | Standardize report headers across all tabs for brand consistency. | Prepare a workbook for distribution with uniform header information.
// AI Prompts: Generate C# code that uses Aspose.Cells to set left, center, and right headers on all worksheets and saves the workbook. | Explain the PageSetup.SetHeader method in Aspose.Cells and list available placeholder codes for dynamic header content.

using System;
using Aspose.Cells;

namespace BatchHeaderSetter
{
    // Load a workbook, define left, center, and right header scripts (file name, page number, date), loop through each worksheet, set the headers with PageSetup.SetHeader, and save the updated file.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (lifecycle rule: use constructor with file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Define the header scripts that will be applied to every worksheet
            string leftHeader = "&F";               // File name
            string centerHeader = "Page &P of &N";   // Page number of total pages
            string rightHeader = "&D";               // Current date

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the PageSetup object of the current worksheet
                PageSetup pageSetup = sheet.PageSetup;

                // Apply the header scripts to the three sections (lifecycle rule: use SetHeader method)
                pageSetup.SetHeader(0, leftHeader);   // Left section
                pageSetup.SetHeader(1, centerHeader); // Center section
                pageSetup.SetHeader(2, rightHeader);  // Right section
            }

            // Save the modified workbook (lifecycle rule: use Save method)
            workbook.Save("output.xlsx");
        }
    }
}
