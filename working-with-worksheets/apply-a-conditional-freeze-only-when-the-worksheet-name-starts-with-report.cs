// Title: Conditional Freeze Panes on Worksheets Starting with “Report” – Aspose.Cells for .NET
// Description: C# example that creates a workbook, adds three sheets (two named with the "Report" prefix), loops through all worksheets, and applies FreezePanes at cell C3 (3 rows × 3 columns) only to sheets whose names begin with "Report" before saving the file.
// Keywords: Aspose.Cells | .NET | C# | freeze panes | conditional freeze | worksheet name prefix | FreezePanes method | Excel automation | Report sheet | Excel workbook generation
// Common Searches: Aspose.Cells freeze panes based on sheet name | How to apply FreezePanes only to worksheets starting with Report | Conditional FreezePanes C# Aspose.Cells | Freeze rows and columns on specific Excel sheets using Aspose | Apply FreezePanes to multiple sheets in .NET
// Developer Intent: Apply FreezePanes to every worksheet whose name starts with "Report" while leaving other sheets unchanged.
// Use Cases: Automated monthly reports where each report tab needs header rows and columns frozen for quick navigation. | Exported data workbooks where only the report sheets receive frozen panes, keeping data sheets fully scrollable. | Template generation that automatically adds a predefined freeze layout to any sheet prefixed with "Report".
// AI Prompts: Generate Aspose.Cells C# code to freeze panes at D4 on all worksheets whose name contains "Summary". | Create a reusable method that accepts a Workbook, a name prefix, a cell address, and row/column counts, then applies FreezePanes to matching sheets. | Show how to apply different freeze pane settings to worksheets based on multiple naming patterns (e.g., "Report", "Data", "Summary") using Aspose.Cells.

using Aspose.Cells;
using System;

// C# example that creates a workbook, adds three sheets (two named with the "Report" prefix), loops through all worksheets, and applies FreezePanes at cell C3 (3 rows × 3 columns) only to sheets whose names begin with "Report" before saving the file.
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
                // Freeze panes at cell C3 with 3 frozen rows and 3 frozen columns
                sheet.FreezePanes("C3", 3, 3);
            }
        }

        // Save the workbook
        workbook.Save("ConditionalFreezeDemo.xlsx");
    }
}
