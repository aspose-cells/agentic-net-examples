using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Tables;

namespace AsposeCellsSlicerDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the table (A1:B5)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["A2"].PutValue("Food");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Drink");
            sheet.Cells["B3"].PutValue(80);
            sheet.Cells["A4"].PutValue("Food");
            sheet.Cells["B4"].PutValue(150);
            sheet.Cells["A5"].PutValue("Snack");
            sheet.Cells["B5"].PutValue(60);

            // Convert the range into a ListObject (table)
            // Add table covering A1:B5 (true indicates the range has headers)
            int tableIndex = sheet.ListObjects.Add("A1", "B5", true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.TableStyleType = TableStyleType.TableStyleMedium2; // optional styling

            // Add a slicer linked to the "Category" column of the table
            // Use zero‑based row and column indices for the upper‑left corner of the slicer
            // Here we place the slicer starting at row 7 (index 6) and column 2 (index 1)
            SlicerCollection slicers = sheet.Slicers;
            int slicerIndex = slicers.Add(table, table.ListColumns[0], 6, 1);
            Slicer slicer = slicers[slicerIndex];

            // Optional: customize slicer appearance
            slicer.Caption = "Category Filter";
            slicer.NumberOfColumns = 1;
            slicer.StyleType = SlicerStyleType.SlicerStyleLight1;
            slicer.WidthPixel = 150;
            slicer.HeightPixel = 120;

            // Save the workbook with the slicer
            workbook.Save("SlicerLinkedToTable.xlsx");
        }
    }
}