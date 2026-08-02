// Title: Add a Left‑Aligned Current Date Footer to All Worksheets with Aspose.Cells for .NET (C#)
// Description: Loads an existing workbook, iterates through every worksheet, sets the left‑aligned footer (section 0) to the current date using the &D placeholder via PageSetup.SetFooter, and saves the updated file.
// Keywords: Aspose.Cells footer date C# | left aligned footer Aspose.Cells | set footer for all worksheets .NET | Excel current date footer programmatically
// Common Searches: Aspose.Cells set left footer date on each sheet | C# add current date footer to all worksheets | PageSetup.SetFooter left section example | how to apply footer to every worksheet Aspose.Cells
// Developer Intent: Add a left‑aligned footer that displays the current date to every worksheet in an existing Excel workbook using Aspose.Cells for .NET.
// Use Cases: Create printable reports where each page shows the generation date in the left footer across all sheets. | Standardize shared workbooks with a consistent date stamp in the left footer before distribution. | Automate batch processing to embed a current‑date footer in multiple workbooks for archival compliance.
// AI Prompts: Generate C# code with Aspose.Cells that adds a left‑aligned current date footer to all worksheets and saves the workbook. | Show how to combine a left‑aligned date footer (&D) and a right‑aligned page number footer (&P of &N) on each sheet using Aspose.Cells. | Explain Aspose.Cells PageSetup footer placeholders (&D, &P, &N) and demonstrate custom footer formatting in C#.

using System;
using Aspose.Cells;

namespace AsposeCellsFooterExample
{
    // Loads an existing workbook, iterates through every worksheet, sets the left‑aligned footer (section 0) to the current date using the &D placeholder via PageSetup.SetFooter, and saves the updated file.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook from file
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Set the left-aligned footer (section 0) to the current date using the &D placeholder
                sheet.PageSetup.SetFooter(0, "&D");
            }

            // Save the modified workbook to a new file
            workbook.Save("output.xlsx");
        }
    }
}
