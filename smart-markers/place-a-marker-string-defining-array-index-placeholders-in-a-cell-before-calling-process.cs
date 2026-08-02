using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerDemo
{
    // Callback implementation to observe smart marker processing
    public class MarkerCallback : ISmartMarkerCallBack
    {
        public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
        {
            // This method is invoked for each smart marker cell before it is populated
            Console.WriteLine($"Callback - Sheet:{sheetIndex}, Row:{rowIndex}, Col:{colIndex}, Table:{tableName}, Column:{columnName}");
        }
    }

    // Simple data class used as a data source
    public class Item
    {
        public string Name { get; set; }
        public int Quantity { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ---------- Place a marker string with array index placeholders ----------
            // The placeholder ${row} will be replaced by the current row index during processing
            // Syntax: &=$Items[${row}].Name  (smart marker for a collection named "Items")
            sheet.Cells["A1"].PutValue("&=$Items[${row}].Name");
            sheet.Cells["B1"].PutValue("&=$Items[${row}].Quantity");

            // ---------- Prepare data source ----------
            List<Item> items = new List<Item>
            {
                new Item { Name = "Apple",  Quantity = 10 },
                new Item { Name = "Banana", Quantity = 20 },
                new Item { Name = "Cherry", Quantity = 30 }
            };

            // ---------- Configure WorkbookDesigner ----------
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook,
                CallBack = new MarkerCallback()   // optional: observe processing
            };

            // Set the collection as a data source; the name must match the smart marker table name ("Items")
            designer.SetDataSource("Items", items);

            // ---------- Process smart markers ----------
            // This will replace the markers with actual data, using the row index placeholder
            designer.Process();

            // ---------- Save the result ----------
            workbook.Save("SmartMarkerWithArrayIndexPlaceholders.xlsx");
        }
    }
}