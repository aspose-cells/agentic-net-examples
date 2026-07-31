// Title: Apply Conditional Freeze Panes to Worksheets Starting with “Report” (Aspose.Cells C#)
// Description: Shows how to generate a workbook, add multiple sheets, and programmatically freeze the top row and left column (cell B2) only on worksheets whose names begin with the prefix “Report”, using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | FreezePanes | conditional freeze | worksheet name prefix | Report sheet | top row left column | B2 freeze | programmatic freeze panes
// Common Searches: Aspose.Cells freeze panes based on sheet name | C# conditional FreezePanes example | Freeze top row and column only on certain worksheets | How to apply FreezePanes to sheets starting with Report | Conditional worksheet freeze in .NET
// Developer Intent: Freeze the first row and column exclusively on sheets whose titles start with “Report”.
// Use Cases: Automated monthly reports where each report tab stays anchored for quick navigation while data tabs scroll freely. | Financial statement exports that lock header rows on report sheets but leave raw data sheets fully scrollable. | Dashboard workbooks that apply a frozen pane to every analysis view prefixed with “Report” to improve readability.
// AI Prompts: Generate C# code with Aspose.Cells that freezes at B2 on every worksheet whose name starts with “Report”, ignoring case. | Provide a .NET example that conditionally calls FreezePanes based on a worksheet name prefix. | Explain how to change the freeze location to C3 for sheets matching the “Report” prefix in the given Aspose.Cells snippet.

using Aspose.Cells;
using System;

// Shows how to generate a workbook, add multiple sheets, and programmatically freeze the top row and left column (cell B2) only on worksheets whose names begin with the prefix “Report”, using Aspose.Cells for .NET.
class ConditionalFreezeDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Prepare sample worksheets
        workbook.Worksheets[0].Name = "Report_January";
        workbook.Worksheets.Add("Data");
        workbook.Worksheets.Add("Report_February");

        // Apply freeze panes only to worksheets whose name starts with "Report"
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            if (sheet.Name.StartsWith("Report", StringComparison.OrdinalIgnoreCase))
            {
                // Freeze the top row and left column (freeze at cell B2)
                sheet.FreezePanes("B2", 1, 1);
            }
        }

        // Save the workbook
        workbook.Save("ConditionalFreezeDemo.xlsx");
    }
}
