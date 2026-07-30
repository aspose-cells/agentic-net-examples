// Title: C# – Prevent Column Drag‑Drop Reordering While Allowing Width Adjustment with Aspose.Cells
// Description: The sample creates a workbook, enables column‑width formatting, turns off sorting to stop drag‑and‑drop reordering, applies full protection without a password, and saves the file as ProtectedWorksheet.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells worksheet protection | C# column width formatting | disable column sorting | prevent column drag drop | protect worksheet without password | AllowFormattingColumn true | AllowSorting false
// Common Searches: Aspose.Cells how to stop column reordering | Allow column width changes but block sorting in .NET | Worksheet protection options for column formatting Aspose.Cells | Protect sheet without password C# Aspose.Cells | Disable column sorting while keeping formatting enabled
// Developer Intent: The developer wants to lock the column order of a worksheet but still let end‑users resize columns.
// Use Cases: Distribute a template where users can adjust column widths for printing, yet the data column sequence must stay fixed. | Generate export files for partners that preserve a predefined column order while permitting visual tweaks to column size. | Create a reporting workbook that protects its layout from accidental reordering but remains flexible for readability adjustments.
// AI Prompts: Provide C# code with Aspose.Cells that protects a worksheet, enables column width changes, and blocks column drag‑and‑drop reordering without using a password. | Explain the effect of setting AllowSorting = false together with ProtectionType.All on user interactions in an Aspose.Cells worksheet. | Show an example of applying worksheet protection that allows formatting columns but prevents sorting, then saving the workbook.

using System;
using Aspose.Cells;

// The sample creates a workbook, enables column‑width formatting, turns off sorting to stop drag‑and‑drop reordering, applies full protection without a password, and saves the file as ProtectedWorksheet.xlsx using Aspose.Cells for .NET.
class WorksheetProtectionDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Access the worksheet's protection settings
        Protection protection = sheet.Protection;

        // Allow users to change column widths
        protection.AllowFormattingColumn = true;

        // Disable column reordering (drag‑and‑drop) by turning off sorting
        protection.AllowSorting = false;

        // Protect the worksheet with all protection options (no password)
        sheet.Protect(ProtectionType.All);

        // Save the protected workbook
        workbook.Save("ProtectedWorksheet.xlsx");
    }
}
