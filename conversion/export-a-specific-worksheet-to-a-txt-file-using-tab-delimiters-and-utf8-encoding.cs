// Title: Export only the active worksheet to a UTF-8 tab-delimited TXT file with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that saves the currently active worksheet of an Aspose.Cells workbook as a UTF-8 tab-delimited text file. | Demonstrate how to configure TxtSaveOptions with a tab separator and UTF-8 encoding to export a single sheet using Aspose.Cells.
// Common Searches: how to export only one sheet to a tab separated txt using Aspose.Cells C# | Aspose.Cells TxtSaveOptions export active worksheet UTF-8 | C# save specific worksheet as tab delimited text file with Aspose.Cells | set encoding and separator for txt export in Aspose.Cells .NET
// Tags: export active worksheet txt Aspose.Cells | tab delimiter option C# Aspose.Cells | UTF-8 encoding TxtSaveOptions .NET | single sheet text conversion Aspose.Cells | configure separator encoding Aspose.Cells

using System;
using System.Text;
using Aspose.Cells;

// // Creates a workbook with two sheets, activates the second sheet, and saves only that sheet as a UTF-8 tab-delimited TXT file using TxtSaveOptions.
class ExportWorksheetToTxt
{
    static void Main()
    {
        // Create a new workbook with two worksheets
        Workbook workbook = new Workbook();
        workbook.Worksheets.Add("Sheet2");

        // Populate data in the first worksheet (index 0)
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Cells["A1"].PutValue("FirstSheet");
        sheet1.Cells["A2"].PutValue(123);

        // Populate data in the second worksheet (index 1) – the one we want to export
        Worksheet sheet2 = workbook.Worksheets[1];
        sheet2.Cells["A1"].PutValue("SecondSheet");
        sheet2.Cells["A2"].PutValue(456);

        // Set the second worksheet as the active sheet so that only it will be exported
        workbook.Worksheets.ActiveSheetIndex = 1;

        // Configure TXT save options: tab delimiter and UTF‑8 encoding
        TxtSaveOptions saveOptions = new TxtSaveOptions
        {
            Separator = '\t',          // Use tab as the column separator
            Encoding = Encoding.UTF8, // Set UTF‑8 encoding
            ExportAllSheets = false   // Export only the active sheet
        };

        // Save the active worksheet to a TXT file
        workbook.Save("SecondSheet.txt", saveOptions);
    }
}
