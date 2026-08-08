// Title: Copy rows between worksheets and preserve row height with Aspose.Cells for .NET (C#)
// Description: This C# example creates a source workbook with rows of different heights, copies the rows and their formatting to a destination workbook using Cells.CopyRows, transfers each row's settings with Row.CopySettings, optionally auto‑fits the rows with AutoFitRows, and saves the result as an Excel file.
// Keywords: Aspose.Cells copy rows | C# copy worksheet rows | preserve row height Aspose.Cells | Row.CopySettings example | AutoFitRows method | Excel row height transfer .NET | copy rows with formatting | Aspose.Cells tutorial
// Common Searches: how to copy rows between worksheets in Aspose.Cells | preserve original row height when copying rows Aspose.Cells C# | auto fit rows after copying data Aspose.Cells | copy row settings including height and style Aspose.Cells | Aspose.Cells copy multiple rows and keep formatting
// Developer Intent: Transfer a block of rows from one worksheet to another while keeping the original row heights and styles, then adjust heights to fit the content.
// Use Cases: Migrate a formatted template section into a report workbook without losing row layout. | Archive selected rows from a live sheet into a separate file while maintaining visual consistency. | Generate a summary sheet by copying rows from user‑filled sheets, preserving visibility, height, and style.
// AI Prompts: Write C# code with Aspose.Cells to copy rows 5‑12 from Sheet1 to Sheet2, keep row heights and styles, then auto‑fit the destination rows. | Explain the difference between Row.CopySettings and AutoFitRows in Aspose.Cells and when to use each. | Provide an Aspose.Cells example that copies rows containing merged cells and ensures the destination rows retain the same height.

using System;
using Aspose.Cells;
using System.Drawing;

// This C# example creates a source workbook with rows of different heights, copies the rows and their formatting to a destination workbook using Cells.CopyRows, transfers each row's settings with Row.CopySettings, optionally auto‑fits the rows with AutoFitRows, and saves the result as an Excel file.
class CopyRowsAndFitDemo
{
    static void Main()
    {
        // Create source workbook and set up sample data with custom row heights
        Workbook srcWb = new Workbook();
        Worksheet srcSheet = srcWb.Worksheets[0];

        srcSheet.Cells["A1"].PutValue("Short text");
        srcSheet.Cells["A2"].PutValue("This is a longer piece of text that will require a taller row.");
        srcSheet.Cells["A3"].PutValue("Another row");

        // Assign different heights to the source rows
        srcSheet.Cells.Rows[0].Height = 15; // default height
        srcSheet.Cells.Rows[1].Height = 30; // taller row
        srcSheet.Cells.Rows[2].Height = 45; // even taller row

        // Create destination workbook
        Workbook destWb = new Workbook();
        Worksheet destSheet = destWb.Worksheets[0];

        // Number of rows to copy (here we know it's 3)
        int rowsToCopy = 3;

        // Copy rows data and formats from source to destination
        destSheet.Cells.CopyRows(srcSheet.Cells, 0, 0, rowsToCopy);

        // Copy row settings (height, style, visibility) from source rows to destination rows
        for (int i = 0; i < rowsToCopy; i++)
        {
            Row srcRow = srcSheet.Cells.Rows[i];
            Row destRow = destSheet.Cells.Rows[i];
            destRow.CopySettings(srcRow, true);
        }

        // Auto‑fit the copied rows in the destination sheet (optional, ensures height matches content)
        destSheet.AutoFitRows(0, rowsToCopy - 1);

        // Save the resulting workbook
        destWb.Save("CopiedRowsAutoFit.xlsx");
    }
}
