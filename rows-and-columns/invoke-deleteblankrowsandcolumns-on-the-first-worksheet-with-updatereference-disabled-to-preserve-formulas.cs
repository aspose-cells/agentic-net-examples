using System;
using Aspose.Cells;

namespace AsposeCellsDeleteBlankRowsAndColumns
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some data with blank rows/columns and formulas
            // Row 1 (index 0) - header
            cells["A1"].PutValue("Item");
            cells["B1"].PutValue("Quantity");
            cells["C1"].PutValue("Price");
            // Row 2 (index 1) - data
            cells["A2"].PutValue("Apple");
            cells["B2"].PutValue(10);
            cells["C2"].PutValue(0.5);
            // Row 3 (index 2) - blank row (will be deleted)
            // Row 4 (index 3) - data with formula referencing previous rows
            cells["A4"].PutValue("Total");
            cells["C4"].Formula = "=SUM(C2:C3)"; // C3 is blank, formula result will be 0.5

            // Add a blank column D (index 3) that will be deleted
            // No data placed in column D

            // Create DeleteOptions with UpdateReference disabled (false) to preserve formulas
            DeleteOptions deleteOptions = new DeleteOptions
            {
                UpdateReference = false
            };

            // Delete blank rows using the options
            cells.DeleteBlankRows(deleteOptions);

            // Delete blank columns using the same options
            cells.DeleteBlankColumns(deleteOptions);

            // Save the workbook to a file
            workbook.Save("DeletedBlankRowsAndColumns.xlsx", SaveFormat.Xlsx);
        }
    }
}