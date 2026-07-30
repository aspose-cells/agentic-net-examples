// Title: Freeze Header Row and Export as Macro‑Enabled XLSM with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds headers and data, freezes the first row using FreezePanes at cell A2, and saves the file as a macro‑enabled XLSM workbook with SaveFormat.Xlsm.
// Keywords: Aspose.Cells freeze panes C# | freeze first row Excel .NET | save workbook as XLSM Aspose | macro enabled Excel export C# | Aspose.Cells SaveFormat.Xlsm
// Common Searches: Aspose.Cells how to freeze header row | C# freeze panes and save as XLSM | Export Aspose.Cells workbook to macro enabled file | Freeze top row in Excel using Aspose.Cells for .NET | SaveFormat.Xlsm example C#
// Developer Intent: Freeze the worksheet’s header row and generate a macro‑enabled XLSM file.
// Use Cases: Produce scroll‑friendly reports where column titles stay visible while navigating large data sets. | Create template workbooks that will later contain VBA macros, keeping the header fixed for user convenience. | Automate data processing pipelines that require macro‑enabled output while preserving header context.
// AI Prompts: Generate C# code to freeze the first two rows and the first column, then save as XLSM using Aspose.Cells. | Explain the differences between SaveFormat.Xls, SaveFormat.Xlsx, and SaveFormat.Xlsm in Aspose.Cells. | Show how to apply both row and column freeze panes before exporting a workbook with macros enabled.

using System;
using Aspose.Cells;

// Creates a workbook, adds headers and data, freezes the first row using FreezePanes at cell A2, and saves the file as a macro‑enabled XLSM workbook with SaveFormat.Xlsm.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample header and data
        worksheet.Cells["A1"].PutValue("Header1");
        worksheet.Cells["B1"].PutValue("Header2");
        worksheet.Cells["A2"].PutValue(100);
        worksheet.Cells["B2"].PutValue(200);

        // Freeze the first row (header) – freeze 1 row, 0 columns at cell A2
        worksheet.FreezePanes("A2", 1, 0);

        // Export the workbook to an XLSM file
        workbook.Save("FrozenHeader.xlsm", SaveFormat.Xlsm);
    }
}
