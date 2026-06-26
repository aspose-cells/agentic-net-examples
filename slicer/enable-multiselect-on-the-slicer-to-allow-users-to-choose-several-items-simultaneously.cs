using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsDemo
{
    class EnableMultiSelectSlicer
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data
                cells["A1"].Value = "Category";
                cells["B1"].Value = "Amount";
                cells["A2"].Value = "Fruit";
                cells["B2"].Value = 100;
                cells["A3"].Value = "Vegetable";
                cells["B3"].Value = 150;
                cells["A4"].Value = "Fruit";
                cells["B4"].Value = 200;
                cells["A5"].Value = "Vegetable";
                cells["B5"].Value = 120;

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D2", "PivotTable1");
                PivotTable pivot = sheet.PivotTables[pivotIndex];
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
                pivot.RefreshData();
                pivot.CalculateData();

                // Enable multiple item selection for the page field (Category)
                if (pivot.PageFields.Count > 0)
                {
                    pivot.PageFields[0].IsMultipleItemSelectionAllowed = true;
                }

                // Add a slicer linked to the Category field
                int slicerIndex = sheet.Slicers.Add(pivot, "F2", "Category");
                Slicer slicer = sheet.Slicers[slicerIndex];
                slicer.StyleType = SlicerStyleType.SlicerStyleLight1;

                // Pre‑select all items to demonstrate multi‑selection
                foreach (var item in slicer.SlicerCache.SlicerCacheItems)
                {
                    item.Selected = true;
                }

                // Save the workbook
                string outputPath = "MultiSelectSlicerDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            EnableMultiSelectSlicer.Run();
        }
    }
}