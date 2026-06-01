using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsNamedRangeSubtotalDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // ---------- Populate sample data ----------
                // Header row
                cells["A1"].PutValue("Category");
                cells["B1"].PutValue("Amount");

                // Data rows
                object[,] data = new object[,]
                {
                    { "A", 100 },
                    { "B", 200 },
                    { "A", 150 },
                    { "B", 250 },
                    { "A", 120 }
                };

                for (int i = 0; i < data.GetLength(0); i++)
                {
                    cells[i + 1, 0].PutValue(data[i, 0]); // Column A
                    cells[i + 1, 1].PutValue(data[i, 1]); // Column B
                }

                // ---------- Apply AutoFilter to the table ----------
                // AutoFilter will be applied to the range A1:B6 (header + data)
                // SetRange(firstRow, firstColumn, totalColumns) – rows are determined automatically
                sheet.AutoFilter.SetRange(0, 0, 2); // columns A and B

                // ---------- Create a named range that refers to the data rows ----------
                // Create a Range object for A2:B6 (data only, without header)
                Aspose.Cells.Range dataRange = cells.CreateRange(1, 0, 5, 2); // firstRow=1, firstColumn=0, totalRows=5, totalColumns=2

                // Add a name to the workbook and set its reference
                int nameIndex = workbook.Worksheets.Names.Add("FilteredData");
                Name namedRange = workbook.Worksheets.Names[nameIndex];
                // RefersTo must be a formula string, e.g., "=Sheet1!A2:B6"
                namedRange.RefersTo = $"={sheet.Name}!{dataRange.RefersTo}";

                // ---------- Add subtotals for the whole table (including header) ----------
                // Define the CellArea covering A1:B6
                CellArea area = CellArea.CreateCellArea(0, 0, 5, 1); // rows 0‑5, columns 0‑1

                // Group by the first column (Category), sum the Amount column, replace existing subtotals,
                // add page breaks between groups, and place summary below each group
                sheet.Cells.Subtotal(
                    area,
                    0,                                 // group by column index 0 (Category)
                    ConsolidationFunction.Sum,         // subtotal function
                    new int[] { 1 },                   // apply to column index 1 (Amount)
                    true,                              // replace existing subtotals
                    true,                              // add page breaks
                    true                               // summary below data
                );

                // ---------- Save the workbook ----------
                string outputPath = "NamedRangeWithSubtotal.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}