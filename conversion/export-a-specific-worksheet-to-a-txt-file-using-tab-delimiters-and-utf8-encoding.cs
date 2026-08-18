// Title: Export a single worksheet to a UTF‑8 tab‑delimited TXT file using Aspose.Cells for .NET
// Description: Demonstrates how to set a worksheet as active, configure TxtSaveOptions with a tab separator and UTF‑8 encoding, and save only that sheet as a TSV‑style text file.
// Keywords: Aspose.Cells export worksheet txt | C# TxtSaveOptions tab delimiter | UTF-8 TSV export Aspose.Cells | .NET save active sheet as text | worksheet to tab delimited file
// Common Searches: Aspose.Cells save active worksheet as txt | C# export sheet to tab delimited file UTF-8 | TxtSaveOptions separator tab example | How to export a single sheet to TSV with Aspose.Cells | Set encoding UTF-8 when saving worksheet to text
// Developer Intent: Save only the chosen worksheet as a UTF‑8 encoded, tab‑separated text file.
// Use Cases: Create a TSV report from a specific sheet for data exchange with third‑party applications. | Generate a UTF‑8 text dump of a worksheet for bulk import into a database or analytics platform. | Automate per‑sheet exports in a scheduled job where each sheet must be delivered as a separate text file.
// AI Prompts: Show C# code that exports a selected worksheet to a UTF‑8 tab‑delimited TXT file with Aspose.Cells, ensuring only that sheet is saved. | Explain how to configure TxtSaveOptions for a custom separator and UTF‑8 encoding when saving a worksheet as text. | Provide step‑by‑step instructions to export multiple worksheets individually to separate TSV files using Aspose.Cells.

using System;
using System.Text;
using Aspose.Cells;

// Demonstrates how to set a worksheet as active, configure TxtSaveOptions with a tab separator and UTF‑8 encoding, and save only that sheet as a TSV‑style text file.
class ExportWorksheetToTxt
{
    static void Main()
    {
        // Create a new workbook and add a second worksheet
        Workbook workbook = new Workbook();
        workbook.Worksheets.Add("SecondSheet");

        // Fill data in the first worksheet
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Cells["A1"].PutValue("Sheet1");
        sheet1.Cells["A2"].PutValue(1);

        // Fill data in the second worksheet (the one we will export)
        Worksheet sheet2 = workbook.Worksheets[1];
        sheet2.Cells["A1"].PutValue("Sheet2");
        sheet2.Cells["A2"].PutValue(2);

        // Make the second worksheet the active sheet so only it will be exported
        workbook.Worksheets.ActiveSheetIndex = 1; // index of "SecondSheet"

        // Configure TxtSaveOptions for tab‑delimited UTF‑8 output
        TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Tsv);
        saveOptions.Separator = '\t';          // tab delimiter
        saveOptions.Encoding = Encoding.UTF8; // UTF‑8 encoding
        // ExportAllSheets remains false (default), so only the active sheet is saved

        // Save the active worksheet to a TXT file
        workbook.Save("SecondSheet.txt", saveOptions);
    }
}
