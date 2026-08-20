// Title: Rename Excel worksheets to the sum of rows and columns using Aspose.Cells for .NET
// Description: Loads a workbook, computes each sheet's data extent by adding its MaxDataRow and MaxDataColumn counts, generates a safe sheet name from the sum, resolves name collisions, renames the worksheets, and saves the file.
// Keywords: Aspose.Cells rename worksheet | C# MaxDataRow MaxDataColumn | Excel sheet naming safe | duplicate sheet name handling | calculate worksheet data extent
// Common Searches: Aspose.Cells rename sheets by data size | C# rename Excel worksheets using MaxDataRow | how to create safe sheet names in Aspose.Cells | prevent duplicate worksheet names Aspose.Cells .NET | sum of rows and columns for sheet name
// Developer Intent: Automatically rename each worksheet to a unique identifier that reflects the total number of populated rows and columns.
// Use Cases: Quickly locate sheets in large workbooks by their data footprint. | Standardize sheet names for automated reporting or data‑pipeline scripts. | Avoid naming conflicts when multiple sheets share the same row‑column total.
// AI Prompts: Generate C# code with Aspose.Cells that renames worksheets based on MaxDataRow + MaxDataColumn and ensures unique, Excel‑compatible names. | Explain the purpose of CellsHelper.CreateSafeSheetName when renaming Excel sheets programmatically. | Show how to append the original sheet index to the new name to guarantee uniqueness.

using System;
using Aspose.Cells;

// Loads a workbook, computes each sheet's data extent by adding its MaxDataRow and MaxDataColumn counts, generates a safe sheet name from the sum, resolves name collisions, renames the worksheets, and saves the file.
class RenameSheetsByDataExtent
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // MaxDataRow and MaxDataColumn are zero‑based; -1 means no data.
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            // Convert to counts (add 1) and treat -1 as 0.
            int rowCount = maxRow >= 0 ? maxRow + 1 : 0;
            int colCount = maxCol >= 0 ? maxCol + 1 : 0;

            // Sum of rows and columns to use as the base name.
            int sum = rowCount + colCount;
            string proposedName = sum.ToString();

            // Ensure the name complies with Excel rules.
            string safeName = CellsHelper.CreateSafeSheetName(proposedName);

            // Resolve possible duplicate names by appending a suffix.
            string finalName = safeName;
            int duplicateIndex = 1;
            while (workbook.Worksheets.Exists(ws => ws != sheet && 
                                                   ws.Name.Equals(finalName, StringComparison.OrdinalIgnoreCase)))
            {
                finalName = $"{safeName}_{duplicateIndex}";
                finalName = CellsHelper.CreateSafeSheetName(finalName);
                duplicateIndex++;
            }

            // Rename the worksheet.
            sheet.Name = finalName;
        }

        // Save the modified workbook.
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
