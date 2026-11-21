// Title: How to load an Excel workbook, bulk insert DataTable rows, and recalculate formulas using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# bulk insert DataTable into an existing worksheet without triggering calculation | disable automatic formula calculation while writing large data sets with Aspose.Cells .NET | recalculate all formulas after programmatically populating a worksheet using Aspose.Cells | load an Excel file, write DataTable rows cell by cell, and save with Aspose.Cells C# example | performance tips for bulk writing cells in Aspose.Cells .NET
// Tags: bulk write cells Aspose.Cells .NET | disable auto calculation Aspose.Cells | recalculate formulas Aspose.Cells | load and save workbook Aspose.Cells | import DataTable to worksheet Aspose.Cells

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// Demonstrates loading 'input.xlsx', creating a DataTable, bulk‑writing its rows into the first worksheet cell‑by‑cell, triggering formula recalculation, and saving the updated workbook as 'output.xlsx' using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file existence
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Prepare sample DataTable
            DataTable table = new DataTable();
            table.Columns.Add("Product", typeof(string));
            table.Columns.Add("Quantity", typeof(int));
            table.Columns.Add("Price", typeof(double));

            for (int i = 1; i <= 1000; i++)
            {
                table.Rows.Add($"Product {i}", i * 2, i * 1.5);
            }

            // Write header
            sheet.Cells[0, 0].PutValue("Product");
            sheet.Cells[0, 1].PutValue("Quantity");
            sheet.Cells[0, 2].PutValue("Price");

            // Bulk write data manually (avoids ImportDataTable which may be unavailable)
            int rowIndex = 1;
            foreach (DataRow dr in table.Rows)
            {
                sheet.Cells[rowIndex, 0].PutValue(dr["Product"]);
                sheet.Cells[rowIndex, 1].PutValue(dr["Quantity"]);
                sheet.Cells[rowIndex, 2].PutValue(dr["Price"]);
                rowIndex++;
            }

            // Recalculate formulas
            workbook.CalculateFormula();

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
