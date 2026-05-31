using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["B1"].PutValue("Product");
            worksheet.Cells["C1"].PutValue("Sales");

            worksheet.Cells["A2"].PutValue("Fruit");
            worksheet.Cells["B2"].PutValue("Apple");
            worksheet.Cells["C2"].PutValue(100);

            worksheet.Cells["A3"].PutValue("Fruit");
            worksheet.Cells["B3"].PutValue("Banana");
            worksheet.Cells["C3"].PutValue(150);

            worksheet.Cells["A4"].PutValue("Vegetable");
            worksheet.Cells["B4"].PutValue("Carrot");
            worksheet.Cells["C4"].PutValue(200);

            // Add a pivot table based on the data range
            int pivotIndex = worksheet.PivotTables.Add("A1:C4", "E3", "PivotTable1");
            PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Refresh and calculate the pivot table to ensure items are generated
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Helper method to log each PivotItem's name and absolute position
            void LogPivotFieldItems(PivotField field)
            {
                Console.WriteLine($"PivotField: {field.Name}");
                if (field.PivotItems != null)
                {
                    foreach (PivotItem item in field.PivotItems)
                    {
                        // Position property gives the absolute position among all pivot items
                        Console.WriteLine($"  Item Name: {item.Name}, Position: {item.Position}");
                    }
                }
                else
                {
                    Console.WriteLine("  No items in this field.");
                }
            }

            // Iterate over all row fields and log their items
            foreach (PivotField rowField in pivotTable.RowFields)
            {
                LogPivotFieldItems(rowField);
            }

            // Iterate over all column fields (if any) and log their items
            foreach (PivotField columnField in pivotTable.ColumnFields)
            {
                LogPivotFieldItems(columnField);
            }

            // Iterate over all data fields and log their items
            foreach (PivotField dataField in pivotTable.DataFields)
            {
                LogPivotFieldItems(dataField);
            }

            // Save the workbook (ensure the directory exists)
            string outputPath = "PivotDebugOutput.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}