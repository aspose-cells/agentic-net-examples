// Title: C# – Insert a Row in Aspose.Cells and Copy Formatting from the Row Above
// Description: Demonstrates how to use Aspose.Cells for .NET to insert a new row with InsertOptions and CopyFormatType.SameAsAbove, automatically inheriting the style of the preceding row, then add data and save the workbook.
// Keywords: Aspose.Cells insert row C# | CopyFormatType SameAsAbove | InsertRows preserve style | Excel row formatting Aspose .NET | programmatic row insertion Aspose.Cells
// Common Searches: Aspose.Cells insert row keep formatting | CopyFormatType SameAsAbove example C# | How to add a row with same style in Aspose.Cells | InsertRows with formatting inheritance Aspose
// Developer Intent: Add a new worksheet row that automatically adopts the formatting of the row above using Aspose.Cells for .NET.
// Use Cases: Appending data to a table while maintaining consistent row appearance. | Expanding a report template where each new row matches existing styling. | Automating Excel generation where inserted rows inherit header or conditional formats.
// AI Prompts: Generate C# code to insert multiple rows in Aspose.Cells, copying the formatting from the previous rows. | Show how to preserve conditional formatting when adding a new row with Aspose.Cells InsertOptions. | Explain the difference between CopyFormatType.SameAsAbove and other copy format options in Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to use Aspose.Cells for .NET to insert a new row with InsertOptions and CopyFormatType.SameAsAbove, automatically inheriting the style of the preceding row, then add data and save the workbook.
class InsertRowWithFormatting
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data and apply formatting to the header row
        cells["A1"].PutValue("Header");
        Style headerStyle = workbook.CreateStyle();
        headerStyle.Font.IsBold = true;
        cells["A1"].SetStyle(headerStyle);

        // Add some data rows
        cells["A2"].PutValue("Item 1");
        cells["A3"].PutValue("Item 2");

        // Prepare insert options to copy formatting from the row above
        InsertOptions insertOptions = new InsertOptions
        {
            CopyFormatType = CopyFormatType.SameAsAbove
        };

        // Insert a new row at index 2 (third row) and copy the formatting from the previous row
        cells.InsertRows(2, 1, insertOptions);

        // Add data to the newly inserted row
        cells["A3"].PutValue("Inserted Item");

        // Save the workbook
        workbook.Save("InsertRowWithFormatting.xlsx");
    }
}
