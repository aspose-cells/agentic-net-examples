// Title: How to copy page‑setup settings from the first worksheet to all other worksheets using Aspose.Cells for .NET (C#)
// AI Prompts: Retrieve the PageSetup object from the first worksheet and assign its properties to each subsequent worksheet in a workbook with Aspose.Cells. | Iterate through a workbook's worksheets and replicate paper size, orientation, margins, and print titles from worksheet index 0 using C#. | Programmatically propagate print area, header/footer margins, and centering settings across all sheets in an Excel file with Aspose.Cells.
// Common Searches: Aspose.Cells copy page setup from first sheet to all sheets C# | C# loop to duplicate worksheet margins and print area using Aspose.Cells | How to apply same page orientation to every worksheet in an Excel workbook with Aspose.Cells | Set identical print titles for all worksheets programmatically Aspose.Cells .NET
// Tags: synchronize worksheet page settings Aspose.Cells | propagate margins and orientation across sheets .NET | duplicate print area for all worksheets C# | apply uniform paper size to Excel workbook Aspose.Cells | set common print titles in multiple sheets C#

using Aspose.Cells;
using System;

// The example loads an Excel workbook, extracts the PageSetup from the first worksheet, loops through the remaining worksheets, and copies key page‑setup properties—including paper size, orientation, margins, header/footer margins, centering options, print area, and title rows/columns—to each sheet before saving the updated file.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the page‑setup settings from the first worksheet
        PageSetup sourceSetup = workbook.Worksheets[0].PageSetup;

        // Loop through all remaining worksheets and copy the settings
        for (int i = 1; i < workbook.Worksheets.Count; i++)
        {
            PageSetup targetSetup = workbook.Worksheets[i].PageSetup;

            // Copy common page‑setup properties
            targetSetup.PaperSize = sourceSetup.PaperSize;
            targetSetup.Orientation = sourceSetup.Orientation;
            targetSetup.BottomMargin = sourceSetup.BottomMargin;
            targetSetup.TopMargin = sourceSetup.TopMargin;
            targetSetup.LeftMargin = sourceSetup.LeftMargin;
            targetSetup.RightMargin = sourceSetup.RightMargin;
            targetSetup.HeaderMargin = sourceSetup.HeaderMargin;
            targetSetup.FooterMargin = sourceSetup.FooterMargin;
            targetSetup.CenterHorizontally = sourceSetup.CenterHorizontally;
            targetSetup.CenterVertically = sourceSetup.CenterVertically;
            targetSetup.PrintArea = sourceSetup.PrintArea;
            targetSetup.PrintTitleRows = sourceSetup.PrintTitleRows;
            targetSetup.PrintTitleColumns = sourceSetup.PrintTitleColumns;
            // Add additional properties here if needed
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
