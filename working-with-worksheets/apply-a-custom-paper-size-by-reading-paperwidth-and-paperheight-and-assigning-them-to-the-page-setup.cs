// Title: Set a Custom Paper Size from Cell Values with Aspose.Cells for .NET (C#)
// Description: C# code that reads width and height values stored in worksheet cells, applies them as a custom paper size via Aspose.Cells PageSetup, and saves the workbook.
// Keywords: Aspose.Cells | C# | CustomPaperSize | PageSetup | custom paper size | read cell values | Excel printing | worksheet page setup | custom page dimensions | Aspose.Cells .NET
// Common Searches: Aspose.Cells set custom paper size from cells | C# Aspose.Cells custom page dimensions example | How to use CustomPaperSize method in Aspose.Cells | Read width and height from Excel and apply page setup | Aspose.Cells PageSetup custom size tutorial
// Developer Intent: Apply a worksheet's PageSetup to a custom size using width and height values read from cells.
// Use Cases: Allow users to specify desired page dimensions directly in an Excel template, then generate a printable file with those exact measurements. | Automate batch conversion of workbooks where each file contains its own custom paper size defined in specific cells. | Create dynamic reports that adapt their print layout based on dimensions entered by the end‑user before export.
// AI Prompts: Generate C# code that reads paper width from cell B1 and height from B2, then sets a custom page size with Aspose.Cells. | Explain the units required by the CustomPaperSize method in Aspose.Cells and how they relate to Excel's measurement system. | Provide a robust example that validates width and height values before calling CustomPaperSize to avoid runtime errors.

using System;
using Aspose.Cells;

// C# code that reads width and height values stored in worksheet cells, applies them as a custom paper size via Aspose.Cells PageSetup, and saves the workbook.
class CustomPaperSizeDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Example: store custom dimensions (in inches) in cells A1 (width) and A2 (height)
        sheet.Cells["A1"].PutValue(5.0);   // width
        sheet.Cells["A2"].PutValue(7.0);   // height

        // Read the custom dimensions from the worksheet
        double width = sheet.Cells["A1"].DoubleValue;
        double height = sheet.Cells["A2"].DoubleValue;

        // Apply custom paper size to the worksheet's page setup
        sheet.PageSetup.PaperSize = PaperSizeType.Custom;
        sheet.PageSetup.CustomPaperSize(width, height);

        // Save the workbook with the custom page setup applied
        workbook.Save("CustomPaperSizeResult.xlsx");
    }
}
