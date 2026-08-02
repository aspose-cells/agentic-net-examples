using System;
using Aspose.Cells;

namespace AsposeCellsInsertRowWithFormatting
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some sample data and formatting in the first three rows
            cells["A1"].PutValue("Header");
            cells["A1"].GetStyle().Font.IsBold = true; // make header bold

            cells["A2"].PutValue("Item 1");
            cells["A3"].PutValue("Item 2");

            // Define insert options to copy formatting from the row above
            InsertOptions insertOptions = new InsertOptions
            {
                CopyFormatType = CopyFormatType.SameAsAbove,
                UpdateReference = true
            };

            // Insert a new row at index 2 (between Item 1 and Item 2)
            // This will copy the formatting from row 1 (the row above the insertion point)
            worksheet.Cells.InsertRows(2, 1, insertOptions);

            // Add data to the newly inserted row
            cells["A3"].PutValue("Inserted Item");

            // Save the workbook
            workbook.Save("InsertRowWithFormatting.xlsx");
        }
    }
}