// Title: Apply conditional FreezePanes to worksheets whose names start with "Report" using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# freeze panes only on sheets whose name starts with Report | how to use FreezePanes conditionally based on worksheet name in .NET | freeze first row and column for selected worksheets with Aspose.Cells | iterate workbook worksheets and apply FreezePanes when name begins with Report | apply conditional FreezePanes to Excel sheets using Aspose.Cells .NET
// Tags: conditional FreezePanes Aspose.Cells | worksheet name filter FreezePanes C# | freeze first row column specific sheets | Aspose.Cells iterate worksheets | apply FreezePanes Excel .NET

using Aspose.Cells;
using System;

// The example loads a workbook, loops through each worksheet, checks if the sheet name starts with "Report" (case‑insensitive), applies FreezePanes(1,1,1,1) to freeze the first row and column on matching sheets, and saves the modified workbook.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate over all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Apply freeze panes only if the worksheet name starts with "Report"
            if (sheet.Name.StartsWith("Report", StringComparison.OrdinalIgnoreCase))
            {
                // Example: freeze the first row and first column
                // Parameters: splitRow, splitColumn, leftColumn, topRow
                sheet.FreezePanes(1, 1, 1, 1);
            }
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
