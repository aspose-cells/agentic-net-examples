using System;
using Aspose.Cells;

namespace AsposeCellsSubtotalTopDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (Category and Value)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["A3"].PutValue("A");
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["A4"].PutValue("B");
            sheet.Cells["B4"].PutValue(200);
            sheet.Cells["A5"].PutValue("B");
            sheet.Cells["B5"].PutValue(250);
            sheet.Cells["A6"].PutValue("C");
            sheet.Cells["B6"].PutValue(300);
            sheet.Cells["A7"].PutValue("C");
            sheet.Cells["B7"].PutValue(350);

            // Define the range that contains the data (including headers)
            CellArea area = CellArea.CreateCellArea("A1", "B7");

            // Apply subtotals:
            // - Group by the first column (Category) => groupBy = 0
            // - Use SUM function for subtotals
            // - Subtotal the second column (Value) => totalList = new int[] { 1 }
            // - Replace existing subtotals: true
            // - Insert page breaks between groups: false (optional)
            // - SummaryBelowData: false (places summary rows above the group)
            sheet.Cells.Subtotal(
                area,
                0,
                ConsolidationFunction.Sum,
                new int[] { 1 },
                true,
                false,
                false);

            // Ensure the outline setting also reflects summary rows at the top
            sheet.Outline.SummaryRowBelow = false;

            // Save the workbook
            workbook.Save("SubtotalTopDemo.xlsx");
        }
    }
}