using System;
using Aspose.Cells;

namespace AsposeCellsAutoFilterDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data with a header row
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Category");
            cells["C1"].PutValue("Price");

            cells["A2"].PutValue("Apple");
            cells["B2"].PutValue("Fruit");
            cells["C2"].PutValue(2.5);

            cells["A3"].PutValue("Carrot");
            cells["B3"].PutValue("Vegetable");
            cells["C3"].PutValue(1.2);

            cells["A4"].PutValue("Banana");
            cells["B4"].PutValue("Fruit");
            cells["C4"].PutValue(3.1);

            cells["A5"].PutValue("Broccoli");
            cells["B5"].PutValue("Vegetable");
            cells["C5"].PutValue(1.8);

            // Apply an AutoFilter to the header range (A1:C5)
            sheet.AutoFilter.Range = "A1:C5";

            // Filter the "Category" column (index 1) to show only "Fruit"
            sheet.AutoFilter.Filter(1, "Fruit");

            // Apply the filter (hide non‑matching rows)
            sheet.AutoFilter.Refresh();

            // Save the workbook in XLSX format (lifecycle: save)
            workbook.Save("AutoFilterDemo.xlsx");

            // Optional: inform the user
            Console.WriteLine("Workbook with AutoFilter saved as AutoFilterDemo.xlsx");
        }
    }
}