// Title: Create flat subtotal rows without outline in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add subtotal rows without outline | generate flat summary rows in Excel with Aspose.Cells .NET | manual subtotal formulas for grouped data using Aspose.Cells | disable outline grouping when inserting subtotals with Aspose.Cells | how to create grand total row in Aspose.Cells workbook
// Tags: flat subtotal rows Aspose.Cells | outline-less subtotals .NET | custom subtotal formulas Aspose.Cells | Excel grand total generation Aspose.Cells | Aspose.Cells workbook subtotal example

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample creates a new workbook, populates sample data, inserts manual subtotal rows with SUM formulas for each category and a grand total row, and saves the file as an .xlsx without using outline grouping.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data
                cells["A1"].PutValue("Category");
                cells["B1"].PutValue("Amount");
                cells["A2"].PutValue("A");
                cells["B2"].PutValue(10);
                cells["A3"].PutValue("A");
                cells["B3"].PutValue(20);
                cells["A4"].PutValue("B");
                cells["B4"].PutValue(30);
                cells["A5"].PutValue("B");
                cells["B5"].PutValue(40);

                // Insert manual subtotal rows (since SubtotalType enum may be unavailable in some versions)
                // Subtotal for Category A
                cells["A6"].PutValue("Subtotal A");
                cells["B6"].Formula = "SUM(B2:B3)";

                // Subtotal for Category B
                cells["A7"].PutValue("Subtotal B");
                cells["B7"].Formula = "SUM(B4:B5)";

                // Overall total
                cells["A8"].PutValue("Grand Total");
                cells["B8"].Formula = "SUM(B2:B5)";

                // Save the workbook
                string outputPath = "SubtotalExample.xlsx";

                // Ensure the directory exists (if a directory part is present)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
