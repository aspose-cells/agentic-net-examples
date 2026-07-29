// Title: Aspose.Cells .NET – Set DeleteOptions.UpdateReference before DeleteBlankRows and DeleteBlankColumns
// Description: C# example that creates a workbook, inserts data with empty rows and columns, configures DeleteOptions.UpdateReference = true, then calls worksheet.Cells.DeleteBlankRows and worksheet.Cells.DeleteBlankColumns using the same options, and finally saves the file as Result.xlsx. Demonstrates the required configuration to keep formulas and references intact.
// Keywords: Aspose.Cells | .NET | C# | DeleteBlankRows | DeleteBlankColumns | DeleteOptions | UpdateReference | remove empty rows | remove empty columns | preserve formulas | cell reference update
// Common Searches: Aspose.Cells DeleteBlankRows keep formulas | How to use DeleteOptions.UpdateReference in .NET | DeleteBlankColumns without breaking references Aspose.Cells | C# example for deleting blank rows and columns | UpdateReference option for DeleteBlankRowsAndColumns
// Developer Intent: Configure DeleteOptions.UpdateReference = true before each DeleteBlankRows or DeleteBlankColumns call so that all cell references, formulas, named ranges, and external links are automatically adjusted.
// Use Cases: Clean up worksheets by removing empty rows while ensuring dependent formulas remain correct. | Eliminate blank columns without disrupting named ranges or cross‑sheet references. | Automate workbook sanitization across multiple sheets in a .NET application, preserving data integrity after deletions.
// AI Prompts: Write C# code using Aspose.Cells that deletes blank rows and columns and updates all formulas and references. | Explain the impact of DeleteOptions.UpdateReference on formulas when calling DeleteBlankRows and DeleteBlankColumns. | Show how to apply DeleteOptions with UpdateReference for each worksheet in a workbook before removing empty rows and columns.

using Aspose.Cells;
using System;

// C# example that creates a workbook, inserts data with empty rows and columns, configures DeleteOptions.UpdateReference = true, then calls worksheet.Cells.DeleteBlankRows and worksheet.Cells.DeleteBlankColumns using the same options, and finally saves the file as Result.xlsx. Demonstrates the required configuration to keep formulas and references intact.
class DeleteBlankRowsColumnsDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate the worksheet with data that includes blank rows and columns
        cells["A1"].PutValue("Header");
        cells["A2"].PutValue("");          // Blank row (row 2)
        cells["A3"].PutValue("Data");
        cells["B1"].PutValue("");          // Blank column (column B)
        cells["C1"].PutValue("ColC");
        cells["C2"].PutValue(123);

        // Configure DeleteOptions to update references in other worksheets
        DeleteOptions deleteOptions = new DeleteOptions
        {
            UpdateReference = true
        };

        // Delete blank rows using the configured options
        worksheet.Cells.DeleteBlankRows(deleteOptions);

        // Delete blank columns using the same options
        worksheet.Cells.DeleteBlankColumns(deleteOptions);

        // Save the modified workbook
        workbook.Save("Result.xlsx", SaveFormat.Xlsx);
    }
}
