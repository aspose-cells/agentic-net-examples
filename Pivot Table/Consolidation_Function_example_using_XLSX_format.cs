using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class Program
    {
        static void Main()
        {
            Run();
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the default worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample data
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");

            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(30);

            // Define the cell area for subtotals (rows 2‑4, columns A‑B)
            CellArea cellArea = new CellArea
            {
                StartRow = 1,
                StartColumn = 0,
                EndRow = 3,
                EndColumn = 1
            };

            // Apply subtotals using different consolidation functions
            worksheet.Cells.Subtotal(cellArea, 0, ConsolidationFunction.Sum, new int[] { 1 });
            worksheet.Cells.Subtotal(cellArea, 0, ConsolidationFunction.Average, new int[] { 1 });
            worksheet.Cells.Subtotal(cellArea, 0, ConsolidationFunction.Max, new int[] { 1 });
            worksheet.Cells.Subtotal(cellArea, 0, ConsolidationFunction.Min, new int[] { 1 });

            // Save the workbook in XLSX format
            workbook.Save("ConsolidationFunctionDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}