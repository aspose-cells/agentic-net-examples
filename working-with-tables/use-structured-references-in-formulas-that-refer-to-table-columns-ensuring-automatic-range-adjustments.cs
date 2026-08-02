// Title: C# – Add a Calculated Column with Structured Reference Shared Formula in an Aspose.Cells Table
// Description: Creates a workbook, defines a ListObject (Excel table) over A1:C4, inserts a "Total" column, and applies a shared structured‑reference formula ([@Column1]+[@Column2]) to the new column. The formula auto‑adjusts as rows are added, the workbook is calculated and saved as StructuredReferenceDemo.xlsx.
// Keywords: Aspose.Cells C# | SetSharedFormula | structured reference | Excel table formula | ListObject example | dynamic column total | shared formula Aspose | auto‑adjust range | C# Excel automation | global finance reporting
// Common Searches: how to use structured references with Aspose.Cells | apply shared formula to a table column in C# | Aspose.Cells SetSharedFormula ListObject example | add calculated column to Excel table using Aspose | auto‑expanding formula in Aspose.Cells table
// Developer Intent: Create an Excel table, add a new column, and set a shared structured‑reference formula that automatically expands with the table size.
// Use Cases: Financial statements where each row’s subtotal is derived from other columns. | Inventory sheets that recalculate totals instantly when new rows are appended. | Template workbooks that use shared formulas to reduce file size and improve performance.
// AI Prompts: Generate C# code that adds a calculated column to an Aspose.Cells ListObject using a structured reference shared formula. | Modify the example to include a third column in the formula (e.g., [@Column1]+[@Column2]+[@Column3]) and recalculate the workbook. | Explain how SetSharedFormula works with tables and how to handle tables that grow beyond the initial row count.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace StructuredReferenceDemo
{
    // Creates a workbook, defines a ListObject (Excel table) over A1:C4, inserts a "Total" column, and applies a shared structured‑reference formula ([@Column1]+[@Column2]) to the new column. The formula auto‑adjusts as rows are added, the workbook is calculated and saved as StructuredReferenceDemo.xlsx.
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

                // ----- Populate sample data -----
                // Header row
                cells["A1"].PutValue("Column1");
                cells["B1"].PutValue("Column2");
                cells["C1"].PutValue("Column3");
                // Data rows
                cells["A2"].PutValue(10);
                cells["B2"].PutValue(20);
                cells["C2"].PutValue(30);

                cells["A3"].PutValue(40);
                cells["B3"].PutValue(50);
                cells["C3"].PutValue(60);

                cells["A4"].PutValue(70);
                cells["B4"].PutValue(80);
                cells["C4"].PutValue(90);

                // ----- Create a table (ListObject) covering the data range -----
                // The table will include the header row and the three data rows (A1:C4)
                int tableIndex = sheet.ListObjects.Add("A1", "C4", true);
                ListObject table = sheet.ListObjects[tableIndex];
                // Give the table a display name (optional, default is Table1)
                table.DisplayName = "MyTable";

                // ----- Add a new column for the calculated total -----
                // Header for the new column
                cells["D1"].PutValue("Total");

                // Set a shared formula for the "Total" column using structured references.
                // The formula adds the values of Column1 and Column2 for each row.
                Cell firstTotalCell = cells["D2"]; // First data cell in the new column
                string structuredFormula = "=[@Column1]+[@Column2]";

                // Apply the shared formula to the range D2:D4 (3 data rows)
                firstTotalCell.SetSharedFormula(structuredFormula, 3, 1);

                // ----- Calculate formulas and save the workbook -----
                workbook.CalculateFormula();

                string outputPath = "StructuredReferenceDemo.xlsx";

                // Ensure the directory exists (handle cases where outputPath has no directory part)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
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
