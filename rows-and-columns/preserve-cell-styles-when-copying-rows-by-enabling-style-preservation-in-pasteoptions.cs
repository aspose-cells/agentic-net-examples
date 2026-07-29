// Title: Copy rows with full formatting using PasteOptions in Aspose.Cells for .NET
// Description: Demonstrates how to copy a styled row from one workbook to another in C# by using Worksheet.Cells.CopyRows together with PasteOptions.PasteType.All, ensuring that values, formulas, colors, fonts and other formats are retained.
// Keywords: Aspose.Cells CopyRows | PasteOptions Preserve Styles | C# copy row formatting | PasteType.All example | retain cell formatting Aspose.Cells | .NET spreadsheet style copy
// Common Searches: Aspose.Cells copy row keep formatting | PasteOptions preserve styles when copying rows | C# copy rows between workbooks with formatting | How to retain background color and bold font in Aspose.Cells | CopyRows with PasteType.All .NET
// Developer Intent: Copy a row from a source worksheet to a destination worksheet while preserving all cell formatting, formulas, and values.
// Use Cases: Migrate a styled header row from a template file to a generated report without losing branding colors. | Duplicate data rows across multiple sheets while keeping conditional formatting and font styles consistent. | Create a new workbook that reuses previously formatted rows to maintain visual standards across documents.
// AI Prompts: Generate C# code that copies multiple rows with their complete formatting from one worksheet to another using Aspose.Cells, including error handling. | Show how to copy rows while preserving only specific style attributes (e.g., background color) with PasteOptions in Aspose.Cells for .NET. | Provide an example of using CopyRows with custom CopyOptions and PasteOptions to transfer rows between workbooks and retain formulas, values, and all formatting.

using System;
using Aspose.Cells;
using System.Drawing;

// Demonstrates how to copy a styled row from one workbook to another in C# by using Worksheet.Cells.CopyRows together with PasteOptions.PasteType.All, ensuring that values, formulas, colors, fonts and other formats are retained.
class PreserveRowStylesExample
{
    static void Main()
    {
        // Create source workbook and add data with a styled row
        Workbook srcWorkbook = new Workbook();
        Worksheet srcSheet = srcWorkbook.Worksheets[0];

        // Populate cells in the first row
        srcSheet.Cells["A1"].PutValue("Header");
        srcSheet.Cells["B1"].PutValue(123);
        srcSheet.Cells["C1"].PutValue(DateTime.Now);

        // Define a style (background color and bold font)
        Style rowStyle = srcWorkbook.CreateStyle();
        rowStyle.ForegroundColor = Color.LightBlue;
        rowStyle.Pattern = BackgroundType.Solid;
        rowStyle.Font.IsBold = true;

        // Apply the style to the entire first row
        srcSheet.Cells.Rows[0].ApplyStyle(rowStyle, new StyleFlag { All = true });

        // Create destination workbook where the row will be copied
        Workbook destWorkbook = new Workbook();
        Worksheet destSheet = destWorkbook.Worksheets[0];

        // Default copy options (no special settings)
        CopyOptions copyOptions = new CopyOptions();

        // Paste options configured to preserve all formats (including styles)
        PasteOptions pasteOptions = new PasteOptions
        {
            PasteType = PasteType.All   // copies values, formulas, formats, etc.
        };

        // Copy the first row from source to destination row index 1 (second row)
        destSheet.Cells.CopyRows(
            srcSheet.Cells,   // source cells
            0,                // source row index
            1,                // destination row index
            1,                // number of rows to copy
            copyOptions,
            pasteOptions);

        // Save workbooks to verify the result
        srcWorkbook.Save("Source.xlsx");
        destWorkbook.Save("Destination.xlsx");
    }
}
