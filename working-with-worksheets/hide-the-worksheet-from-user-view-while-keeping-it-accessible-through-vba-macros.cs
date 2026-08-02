// Title: Hide a Worksheet with VisibilityType.VeryHidden while Keeping VBA Access in Aspose.Cells for .NET (C#)
// Description: Shows how to create a macro‑enabled .xlsm workbook, add a sheet, set its VisibilityType to VeryHidden so it is invisible in the Excel UI yet reachable by VBA macros, and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | VeryHidden | hide worksheet | macro enabled workbook | XLSM | VBA access | VisibilityType | Excel security | programmatic sheet hiding
// Common Searches: Aspose.Cells hide worksheet VeryHidden C# | Create macro enabled Excel file with hidden sheet using Aspose.Cells | VeryHidden sheet accessible by VBA Aspose.Cells | Set VisibilityType.VeryHidden in .NET | Hide Excel sheet from UI but use in macro Aspose
// Developer Intent: Programmatically hide a worksheet from the Excel interface while still allowing VBA macros to read or modify its data.
// Use Cases: Store confidential configuration values that macros can retrieve at runtime. | Hide lookup tables or intermediate results used by VBA functions. | Protect proprietary formulas by placing them on a VeryHidden sheet accessible only through code.
// AI Prompts: Generate C# code with Aspose.Cells that adds a worksheet, sets VisibilityType.VeryHidden, enables macros, and saves the workbook as an .xlsm file. | Provide a VBA macro example that reads cell A1 from a VeryHidden worksheet created by Aspose.Cells. | Explain how to programmatically unhide a VeryHidden sheet via VBA after the workbook is opened.

using System;
using Aspose.Cells;

// Shows how to create a macro‑enabled .xlsm workbook, add a sheet, set its VisibilityType to VeryHidden so it is invisible in the Excel UI yet reachable by VBA macros, and save the file using Aspose.Cells for .NET.
class HideWorksheetDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a new worksheet that will be hidden
            Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenData");

            // Optional: put some data in the hidden sheet
            hiddenSheet.Cells["A1"].PutValue("Secret Value");

            // Hide the worksheet using VeryHidden so it cannot be shown via the UI,
            // but it remains accessible to VBA macros
            hiddenSheet.VisibilityType = VisibilityType.VeryHidden;

            // Enable macros in the workbook (required for macro-enabled files)
            workbook.Settings.EnableMacros = true;

            // Save the workbook as a macro‑enabled file
            workbook.Save("HiddenSheetDemo.xlsm", SaveFormat.Xlsm);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
