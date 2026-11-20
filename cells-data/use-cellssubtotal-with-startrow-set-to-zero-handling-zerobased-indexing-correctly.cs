// Title: C# example: calculate Excel subtotal with Aspose.Cells using zero‑based startRow indexing
// Common Searches: Aspose.Cells calculate subtotal starting from first row C# zero based index | how to sum a column with Aspose.Cells when startRow is zero | C# Aspose.Cells manual subtotal using zero‑based row numbers | Excel subtotal with Aspose.Cells API zero based row indexing example | Aspose.Cells workbook subtotal row index zero C# code sample
// Tags: Aspose.Cells manual subtotal zero based indexing | C# sum column values Aspose.Cells | Excel subtotal calculation Aspose.Cells C# | zero based row indexing Aspose.Cells | Aspose.Cells generate total row workbook

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new workbook, fills columns A and B with category and amount data, manually computes the sum of the Amount column using zero‑based row indices, writes a "Total" label and the calculated sum to a new row, and saves the file as SubtotalResult.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["A2"].PutValue("Food");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Food");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("Transport");
            sheet.Cells["B4"].PutValue(15);
            sheet.Cells["A5"].PutValue("Transport");
            sheet.Cells["B5"].PutValue(25);

            // Manually add a total row (since Subtotal API may not be available)
            int totalRowIndex = 6; // Row 7 (zero‑based index)
            sheet.Cells[totalRowIndex, 0].PutValue("Total");
            // Sum the Amount column (B) from rows 2 to 5 (zero‑based indices 1‑4)
            double sum = 0;
            for (int r = 1; r <= 4; r++)
            {
                var val = sheet.Cells[r, 1].Value;
                if (val is double d) sum += d;
                else if (val is int i) sum += i;
                else if (double.TryParse(val?.ToString(), out double parsed)) sum += parsed;
            }
            sheet.Cells[totalRowIndex, 1].PutValue(sum);

            // Define output file path
            string outputPath = "SubtotalResult.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
