// Title: Freeze Header Row in Aspose.Cells (C#) and Save as Macro‑Enabled XLSM
// Description: Creates a new workbook, writes three header cells, freezes the first row with FreezePanes, and saves the result as an XLSM macro‑enabled file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | FreezePanes | freeze header row | macro enabled workbook | XLSM export | worksheet freeze | save as xlsm | Aspose.Cells tutorial
// Common Searches: Aspose.Cells freeze first row C# | save workbook as XLSM Aspose.Cells | how to freeze panes in Aspose.Cells .NET | export macro‑enabled file with Aspose.Cells | C# freeze header and export XLSM
// Developer Intent: Freeze the worksheet’s header row and export the workbook as a macro‑enabled XLSM file.
// Use Cases: Generate a scrolling report where column headers stay visible and deliver it as an XLSM file. | Provide a template that preserves frozen headers while allowing users to add or edit macros. | Automate data exports that require a fixed header row and must be compatible with macro‑enabled Excel.
// AI Prompts: Show how to freeze multiple rows and columns in Aspose.Cells before saving as XLSM. | Give C# code that adds a simple VBA macro after freezing panes and then saves the workbook. | Explain how to calculate the FreezePanes parameters dynamically based on a variable number of header rows.

using System;
using Aspose.Cells;

// Creates a new workbook, writes three header cells, freezes the first row with FreezePanes, and saves the result as an XLSM macro‑enabled file using Aspose.Cells for .NET.
class FreezeAndExportXlsm
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add header values to the first row
        worksheet.Cells["A1"].PutValue("Header1");
        worksheet.Cells["B1"].PutValue("Header2");
        worksheet.Cells["C1"].PutValue("Header3");

        // Freeze the first row (header) so it stays visible while scrolling
        // Parameters: row index, column index, number of frozen rows, number of frozen columns
        // Row index 1 means the freeze line is just below the first row (index 0)
        worksheet.FreezePanes(1, 0, 1, 0);

        // Save the workbook as an XLSM file (macro-enabled workbook)
        workbook.Save("FrozenHeaders.xlsm", SaveFormat.Xlsm);
    }
}
