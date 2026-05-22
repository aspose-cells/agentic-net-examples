using System;
using Aspose.Cells;

namespace DeleteBlankRowsExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate data with some blank rows
            cells["A1"].PutValue("Header");
            cells["A2"].PutValue("Data1");
            // Row 3 will be blank
            cells["A4"].PutValue("Data2");
            // Row 5 will be blank
            cells["A6"].PutValue("Data3");

            // Delete all blank rows using the default DeleteBlankRows method
            // (no DeleteOptions are supplied, so UpdateReference remains false)
            cells.DeleteBlankRows();

            // Save the workbook
            workbook.Save("DeletedBlankRows.xlsx", SaveFormat.Xlsx);
        }
    }
}