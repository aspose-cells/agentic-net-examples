using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotTabularDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // Prepare source data on the first worksheet
            // -------------------------------------------------
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Header
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("SubCategory");
            dataSheet.Cells["C1"].PutValue("Amount");

            // Sample rows
            dataSheet.Cells["A2"].PutValue("Fruit");
            dataSheet.Cells["B2"].PutValue("Apple");
            dataSheet.Cells["C2"].PutValue(120);

            dataSheet.Cells["A3"].PutValue("Fruit");
            dataSheet.Cells["B3"].PutValue("Banana");
            dataSheet.Cells["C3"].PutValue(80);

            dataSheet.Cells["A4"].PutValue("Vegetable");
            dataSheet.Cells["B4"].PutValue("Carrot");
            dataSheet.Cells["C4"].PutValue(50);

            dataSheet.Cells["A5"].PutValue("Vegetable");
            dataSheet.Cells["B5"].PutValue("Potato");
            dataSheet.Cells["C5"].PutValue(70);

            // -------------------------------------------------
            // Create a worksheet that will host the pivot table
            // -------------------------------------------------
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

            // Add a pivot table based on the source range
            int pivotIndex = pivotSheet.PivotTables.Add("=Data!A1:C5", "A3", "PivotTable1");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Add fields: Category and SubCategory as rows, Amount as data
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Row, "SubCategory");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // -------------------------------------------------
            // Set the layout to Tabular form
            // -------------------------------------------------
            pivotTable.ShowInTabularForm();

            // Refresh and calculate to populate the view
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // -------------------------------------------------
            // Verify column alignment (Tabular layout should have
            // each row field displayed without outline merging)
            // -------------------------------------------------
            bool alignmentOk = true;
            foreach (PivotField rowField in pivotTable.RowFields)
            {
                // In Tabular form, ShowInOutlineForm must be false
                if (rowField.ShowInOutlineForm)
                {
                    alignmentOk = false;
                    Console.WriteLine($"Field '{rowField.Name}' is still in outline form.");
                }
            }

            if (alignmentOk)
            {
                Console.WriteLine("PivotTable is correctly displayed in Tabular form (column alignment verified).");
            }
            else
            {
                Console.WriteLine("PivotTable column alignment verification failed.");
            }

            // -------------------------------------------------
            // Save the workbook
            // -------------------------------------------------
            workbook.Save("PivotTableTabularFormDemo.xlsx");
        }
    }
}