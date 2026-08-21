// Title: Hide a Worksheet (VeryHidden) in an XLSM file with Aspose.Cells for .NET
// Description: Demonstrates how to create a macro‑enabled workbook, add a sheet, set its VisibilityType to VeryHidden so it stays invisible in Excel, and save the file as XLSM for later VBA access using Aspose.Cells for .NET.
// Keywords: Aspose.Cells hide worksheet | VeryHidden sheet .NET | macro enabled workbook XLSX | save as XLSM Aspose | Excel VBA hidden sheet | worksheet VisibilityType | C# Aspose.Cells example | global developers | US .NET developers | EU Excel automation
// Common Searches: Aspose.Cells set worksheet VeryHidden | Create XLSM with hidden sheet using C# | How to keep a sheet invisible but accessible to VBA | Hide Excel sheet programmatically Aspose | VeryHidden worksheet example .NET
// Developer Intent: Make a worksheet invisible in the Excel UI while still allowing VBA macros to read or unhide it.
// Use Cases: Store configuration or lookup tables that macros read at runtime. | Protect proprietary formulas by placing them on a VeryHidden sheet. | Distribute template workbooks where end users cannot see internal calculation sheets.
// AI Prompts: Generate VBA code to unhide a VeryHidden worksheet created with Aspose.Cells. | Show C# to add multiple VeryHidden sheets and apply password protection in an XLSM workbook. | Explain how to toggle a worksheet between VeryHidden and Visible using Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to create a macro‑enabled workbook, add a sheet, set its VisibilityType to VeryHidden so it stays invisible in Excel, and save the file as XLSM for later VBA access using Aspose.Cells for .NET.
class HideWorksheetDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a new worksheet that we want to hide
            Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");

            // Set the worksheet to VeryHidden.
            // This makes the sheet invisible in the Excel UI,
            // but it can still be accessed and made visible by VBA macros.
            hiddenSheet.VisibilityType = VisibilityType.VeryHidden;

            // Save the workbook as a macro‑enabled file (XLSM) so that VBA macros can be added later if needed.
            string outputPath = "HiddenSheetDemo.xlsm";
            workbook.Save(outputPath, SaveFormat.Xlsm);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
