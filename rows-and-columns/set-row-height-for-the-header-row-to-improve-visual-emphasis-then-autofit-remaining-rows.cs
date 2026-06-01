using System;
using Aspose.Cells;

namespace AsposeCellsHeaderRowExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate header row (row 0) and some sample data rows
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Description");
            cells["C1"].PutValue("Price");

            cells["A2"].PutValue("Apple");
            cells["B2"].PutValue("Fresh red apple from the orchard");
            cells["C2"].PutValue(1.20);

            cells["A3"].PutValue("Banana");
            cells["B3"].PutValue("Ripe bananas, sweet and soft");
            cells["C3"].PutValue(0.80);

            // Set a custom height for the header row to make it stand out (height in points)
            cells.SetRowHeight(0, 30); // e.g., 30 points

            // Determine the last row that contains data
            int lastDataRow = cells.MaxDataRow; // zero‑based index

            // Auto‑fit all rows except the header (rows 1 … lastDataRow)
            if (lastDataRow >= 1)
            {
                sheet.AutoFitRows(1, lastDataRow);
            }

            // Optionally auto‑fit columns for better visibility
            sheet.AutoFitColumns();

            // Save the workbook
            string outputPath = "HeaderRowHeightDemo.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}