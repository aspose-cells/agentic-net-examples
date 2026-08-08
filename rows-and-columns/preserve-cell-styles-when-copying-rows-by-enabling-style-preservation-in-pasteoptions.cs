// Title: Preserve Cell Styles When Copying Rows with Aspose.Cells for .NET
// Description: Demonstrates how to copy a row from one worksheet to another while retaining its formatting by using CopyRows together with PasteOptions set to PasteType.Formats. The example creates a styled header row, copies it to a new location, copies a data row without formatting, and saves the workbook as PreserveRowStyles.xlsx.
// Keywords: Aspose.Cells | CopyRows | PasteOptions | .NET | C# | preserve formatting | row style copy | Excel automation | style preservation | PasteType.Formats
// Common Searches: Aspose.Cells copy row keep formatting | PasteOptions preserve styles Aspose.Cells .NET | CopyRows with PasteType Formats example | How to retain cell style when copying rows in Aspose.Cells | C# copy header row with formatting using Aspose
// Developer Intent: Copy rows between worksheets while maintaining the original cell formatting.
// Use Cases: Duplicate a formatted header row in a report workbook without losing its style. | Create a template with styled rows and programmatically insert them into generated spreadsheets. | Separate value-only copying from style-preserving copying in the same automation workflow.
// AI Prompts: Write C# code that copies multiple rows with their formatting using Aspose.Cells CopyRows and PasteOptions. | Show how to copy a row with values only, then copy another row preserving its style in Aspose.Cells for .NET. | Explain the effect of PasteType.Formats on the CopyRows method in Aspose.Cells.

using Aspose.Cells;
using System.Drawing;

// Demonstrates how to copy a row from one worksheet to another while retaining its formatting by using CopyRows together with PasteOptions set to PasteType.Formats. The example creates a styled header row, copies it to a new location, copies a data row without formatting, and saves the workbook as PreserveRowStyles.xlsx.
class PreserveRowStyles
{
    static void Main()
    {
        // Create source workbook and apply a style to the first row
        Workbook srcWorkbook = new Workbook();
        Worksheet srcSheet = srcWorkbook.Worksheets[0];

        // Define a style (yellow background, bold font)
        Style headerStyle = srcWorkbook.CreateStyle();
        headerStyle.ForegroundColor = Color.Yellow;
        headerStyle.Pattern = BackgroundType.Solid;
        headerStyle.Font.IsBold = true;

        // Populate header cells and apply the style
        srcSheet.Cells["A1"].PutValue("Header1");
        srcSheet.Cells["B1"].PutValue("Header2");
        srcSheet.Cells["A1"].SetStyle(headerStyle);
        srcSheet.Cells["B1"].SetStyle(headerStyle);

        // Add a data row below the header
        srcSheet.Cells["A2"].PutValue(10);
        srcSheet.Cells["B2"].PutValue(20);

        // Create destination workbook
        Workbook destWorkbook = new Workbook();
        Worksheet destSheet = destWorkbook.Worksheets[0];

        // Prepare copy and paste options
        CopyOptions copyOptions = new CopyOptions(); // default options
        PasteOptions pasteOptions = new PasteOptions
        {
            // Preserve only formatting (styles) when copying rows
            PasteType = PasteType.Formats
        };

        // Copy the header row (row index 0) to destination row index 5,
        // preserving its style via PasteOptions
        destSheet.Cells.CopyRows(
            srcSheet.Cells,      // source cells
            0,                   // source row index
            5,                   // destination row index
            1,                   // number of rows to copy
            copyOptions,        // copy options (default)
            pasteOptions);      // paste options with style preservation

        // Copy the data row without special paste options (values only)
        destSheet.Cells.CopyRows(srcSheet.Cells, 1, 6, 1);

        // Save the result workbook
        destWorkbook.Save("PreserveRowStyles.xlsx");
    }
}
