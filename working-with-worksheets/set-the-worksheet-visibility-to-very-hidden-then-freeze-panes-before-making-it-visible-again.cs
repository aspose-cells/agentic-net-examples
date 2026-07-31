// Title: Aspose.Cells C# – Set Worksheet VeryHidden, Freeze Panes, Then Unhide
// Description: Creates a workbook, adds a temporary visible sheet, marks the first worksheet as VeryHidden, applies FreezePanes at C3 (2 rows × 2 columns), restores the worksheet to Visible, optionally removes the temporary sheet, and saves the file.
// Keywords: Aspose.Cells VeryHidden worksheet | Aspose.Cells FreezePanes C# | unhide VeryHidden sheet Aspose.Cells | temporary worksheet Aspose.Cells | save workbook Aspose.Cells C#
// Common Searches: Aspose.Cells set worksheet VeryHidden then visible | freeze panes on hidden sheet Aspose.Cells | how to keep one sheet visible while hiding others Aspose.Cells | C# Aspose.Cells remove temporary sheet after hiding
// Developer Intent: Hide a worksheet as VeryHidden, freeze panes, and make it visible again while satisfying Aspose.Cells' requirement for at least one visible sheet.
// Use Cases: Protect layout by hiding a sheet VeryHidden, freezing headers, then revealing it for end‑users. | Generate templates where internal sheets stay hidden during creation and are shown before distribution. | Work around Aspose.Cells' visibility rule by using a temporary sheet while manipulating other sheets.
// AI Prompts: Write C# code using Aspose.Cells that sets a worksheet to VeryHidden, freezes panes at a given cell, then restores visibility, ensuring a temporary visible sheet exists. | Provide an Aspose.Cells example that removes the temporary worksheet after the hidden sheet is made visible again.

using System;
using Aspose.Cells;

// Creates a workbook, adds a temporary visible sheet, marks the first worksheet as VeryHidden, applies FreezePanes at C3 (2 rows × 2 columns), restores the worksheet to Visible, optionally removes the temporary sheet, and saves the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (contains one default worksheet)
            Workbook workbook = new Workbook();

            // Add a temporary worksheet so the workbook always has at least one visible sheet
            Worksheet tempSheet = workbook.Worksheets.Add("Temp");

            // Access the first worksheet (the one we will manipulate)
            Worksheet sheet = workbook.Worksheets[0];

            // Hide the worksheet as VeryHidden (allowed because Temp sheet is visible)
            sheet.VisibilityType = VisibilityType.VeryHidden;

            // Freeze panes at cell C3 with 2 frozen rows and 2 frozen columns
            sheet.FreezePanes("C3", 2, 2);

            // Make the worksheet visible again
            sheet.VisibilityType = VisibilityType.Visible;

            // Optional: remove the temporary worksheet
            int tempIndex = workbook.Worksheets.IndexOf(tempSheet);
            if (tempIndex >= 0)
                workbook.Worksheets.RemoveAt(tempIndex);

            // Save the workbook
            workbook.Save("VeryHiddenFreezeDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
