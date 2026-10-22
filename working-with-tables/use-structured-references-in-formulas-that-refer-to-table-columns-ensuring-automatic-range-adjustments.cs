// Title: Add a calculated "Total" column and a grand‑total row to an Excel ListObject using structured references with Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a ListObject, inserts a new column, and applies the formula =[@Quantity]*[@Price] to each data row using structured references. | Show how to enable the totals row of an Aspose.Cells table and assign a SUM formula that aggregates the calculated column. | Demonstrate saving the workbook after adding the table, calculated column, and totals row to an .xlsx file.
// Common Searches: asp.net how to add a calculated column to an Excel table with Aspose.Cells structured reference | c# Aspose.Cells set formula for each row in ListObject using @ syntax | asp.net create totals row in Excel table and sum a column with Aspose.Cells | using Aspose.Cells .NET to automatically expand table range when adding columns
// Tags: Aspose.Cells add calculated column with structured reference | Aspose.Cells ListObject totals row SUM formula | Aspose.Cells dynamic table range expansion .NET | C# Excel table structured reference formula | Aspose.Cells create ListObject from range

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace StructuredReferenceExample
{
    // The example creates a new workbook, defines a ListObject named SalesTable covering A1:C5, adds a calculated "Total" column using the structured reference formula =[@Quantity]*[@Price] for each data row, enables the table's totals row, sets a SUM formula over the Total column, and saves the result as StructuredReferenceDemo.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Access the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Data";

                // Populate sample data (A1:C5).
                // Header row.
                sheet.Cells["A1"].PutValue("Item");
                sheet.Cells["B1"].PutValue("Quantity");
                sheet.Cells["C1"].PutValue("Price");

                // Data rows.
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["C2"].PutValue(0.5);

                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["C3"].PutValue(0.3);

                sheet.Cells["A4"].PutValue("Cherry");
                sheet.Cells["B4"].PutValue(15);
                sheet.Cells["C4"].PutValue(0.8);

                sheet.Cells["A5"].PutValue("Date");
                sheet.Cells["B5"].PutValue(5);
                sheet.Cells["C5"].PutValue(1.2);

                // Define the range that will become a table (A1:C5).
                int firstRow = 0;        // zero‑based index for row 1
                int firstColumn = 0;     // zero‑based index for column A
                int totalRows = 5;
                int totalColumns = 3;

                // Add a ListObject (table) to the worksheet.
                int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn,
                    firstRow + totalRows - 1, firstColumn + totalColumns - 1, true);
                ListObject table = sheet.ListObjects[tableIndex];
                table.DisplayName = "SalesTable"; // Table name used in structured references.

                // Add a new column for "Total" (Quantity * Price) using a structured reference formula.
                int totalColumnIndex = totalColumns; // Next column after existing ones (D column).
                sheet.Cells[0, totalColumnIndex].PutValue("Total"); // Header for the new column.

                // Set the formula for each data row using structured reference.
                string formula = "=[@Quantity]*[@Price]";
                for (int row = 1; row < totalRows; row++) // start from row 2 (index 1)
                {
                    sheet.Cells[row, totalColumnIndex].Formula = formula;
                }

                // Enable the total row and set its formulas.
                table.ShowTotals = true;
                int totalRowIndex = totalRows; // zero‑based index for row 6.
                sheet.Cells[totalRowIndex, 0].PutValue("Grand Total");
                string totalFormula = "=SUM(SalesTable[Total])";
                sheet.Cells[totalRowIndex, totalColumnIndex].Formula = totalFormula;

                // Save the workbook to a file.
                string outputPath = "StructuredReferenceDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
