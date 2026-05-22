using System;
using Aspose.Cells;

namespace DeleteBlankRowsAndColumnsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate the worksheet with sample data that includes blank rows and columns
            cells["A1"].PutValue("Header");
            cells["A2"].PutValue("");               // Blank row (empty string)
            cells["A3"].PutValue("Data");
            cells["B1"].PutValue("");               // Blank column (empty string)
            cells["C1"].PutValue("Extra");

            // -----------------------------------------------------------------
            // Delete blank rows with UpdateReference set to true
            // -----------------------------------------------------------------
            DeleteBlankOptions rowOptions = new DeleteBlankOptions
            {
                // Ensure references in other worksheets are updated after deletion
                UpdateReference = true,
                // Treat empty strings as blanks (optional, shown for completeness)
                EmptyStringAsBlank = true
            };
            // Perform the deletion of blank rows using the configured options
            sheet.Cells.DeleteBlankRows(rowOptions);

            // -----------------------------------------------------------------
            // Delete blank columns with UpdateReference set to true
            // -----------------------------------------------------------------
            DeleteBlankOptions columnOptions = new DeleteBlankOptions
            {
                // Again, make sure references are updated for column deletion
                UpdateReference = true,
                // Treat empty strings as blanks (optional)
                EmptyStringAsBlank = true
            };
            // Perform the deletion of blank columns using the configured options
            sheet.Cells.DeleteBlankColumns(columnOptions);

            // Save the modified workbook
            workbook.Save("DeletedBlankRowsAndColumns.xlsx", SaveFormat.Xlsx);
        }
    }
}