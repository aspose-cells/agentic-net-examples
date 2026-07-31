// Title: Aspose.Cells C# – Hide Zero Values on Worksheets Starting with “Q” via WorksheetCollection
// Description: Load an Excel file, iterate its WorksheetCollection, and set DisplayZeros = false on every sheet whose name begins with "Q" (case‑insensitive). The workbook is then saved with the zero values hidden on the selected worksheets.
// Keywords: Aspose.Cells hide zeros C# | WorksheetCollection iterate | DisplayZeros property | conditional worksheet formatting | hide zero values Excel .NET | filter worksheets by name | case‑insensitive sheet prefix
// Common Searches: Aspose.Cells hide zero values on specific sheets | C# hide zeros on worksheets that start with Q | set DisplayZeros for selected worksheets Aspose | iterate worksheets and apply display settings | how to suppress zero values in Aspose.Cells workbook
// Developer Intent: Programmatically suppress zero values on every worksheet whose name starts with the letter “Q”.
// Use Cases: Quarterly financial reports (Q1, Q2, …) where zero amounts should be invisible. | Dashboard workbooks that use Q‑prefixed tabs and need a cleaner visual layout. | Automated workbook preparation before distribution, removing zero clutter from targeted sheets.
// AI Prompts: Generate C# code using Aspose.Cells to hide zero values on all worksheets whose names start with "Q" and save the result. | Explain the DisplayZeros property and demonstrate how to apply it conditionally based on worksheet names. | Show a case‑insensitive example of filtering worksheets by prefix and disabling zero display in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Load an Excel file, iterate its WorksheetCollection, and set DisplayZeros = false on every sheet whose name begins with "Q" (case‑insensitive). The workbook is then saved with the zero values hidden on the selected worksheets.
class HideZeroValues
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the collection
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // If the worksheet name starts with "Q" (case‑insensitive)
            if (sheet.Name.StartsWith("Q", StringComparison.OrdinalIgnoreCase))
            {
                // Hide zero values on this sheet
                sheet.DisplayZeros = false;
            }
        }

        // Save the workbook with the changes applied
        workbook.Save("output.xlsx");
    }
}
