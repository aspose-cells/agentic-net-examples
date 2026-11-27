// Title: Insert a Year‑Over‑Year Growth % calculated column into an Excel sheet with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add calculated column for YoY growth percentage | How to compute year over year sales increase using SUMIFS in Aspose.Cells | Generate dynamic YoY growth formula in an Excel workbook with .NET | Apply percentage number format to a new column using Aspose.Cells | Create a YoY growth column based on previous year sales in C#
// Tags: Aspose.Cells insert YoY growth field | C# SUMIFS previous year sales formula | Excel percentage style Aspose.Cells | dynamic Excel formula generation .NET | custom worksheet column addition Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace YoYGrowthApp
{
    // The example shows how to add a "YoY Growth %" column to the first worksheet of a workbook, build a SUMIFS formula that fetches the prior year's sales for each product, wrap the calculation with IFERROR, and format the result as a percentage with two decimal places using Aspose.Cells for .NET.
    public class YoYGrowthCalculator
    {
        // Adds a calculated column "YoY Growth %" to the first worksheet.
        public static void AddYoYGrowth(Workbook workbook)
        {
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            int lastRow = cells.MaxDataRow; // zero‑based index

            // Header for the new column (D)
            cells[0, 3].PutValue("YoY Growth %");

            // Build formulas for each data row
            for (int row = 1; row <= lastRow; row++)
            {
                string productCell = cells[row, 0].Name; // e.g., A2
                string yearCell = cells[row, 1].Name;    // e.g., B2
                string salesCell = cells[row, 2].Name;   // e.g., C2

                string sumRange = "$C:$C";
                string productRange = "$A:$A";
                string yearRange = "$B:$B";

                // SUMIFS to get previous year's sales for the same product
                string prevYearSales = $"SUMIFS({sumRange}, {productRange}, {productCell}, {yearRange}, {yearCell}-1)";

                // Complete formula with IFERROR
                string formula = $"IFERROR(({salesCell}-{prevYearSales})/{prevYearSales},0)";

                cells[row, 3].Formula = formula;
            }

            // Format the new column as percentage with two decimal places
            Style style = workbook.CreateStyle();
            style.Number = 10; // 0.00%
            StyleFlag flag = new StyleFlag { NumberFormat = true };
            cells.CreateRange(1, 3, lastRow, 1).ApplyStyle(style, flag);
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "Input.xlsx";
            string outputPath = "Output.xlsx";

            try
            {
                // Ensure input file exists; create a sample if missing
                if (!File.Exists(inputPath))
                {
                    Workbook sample = new Workbook();
                    Worksheet ws = sample.Worksheets[0];
                    Cells c = ws.Cells;

                    // Header
                    c[0, 0].PutValue("Product");
                    c[0, 1].PutValue("Year");
                    c[0, 2].PutValue("Sales");

                    // Sample data
                    c[1, 0].PutValue("A"); c[1, 1].PutValue(2020); c[1, 2].PutValue(100);
                    c[2, 0].PutValue("A"); c[2, 1].PutValue(2021); c[2, 2].PutValue(150);
                    c[3, 0].PutValue("B"); c[3, 1].PutValue(2020); c[3, 2].PutValue(200);
                    c[4, 0].PutValue("B"); c[4, 1].PutValue(2021); c[4, 2].PutValue(250);

                    sample.Save(inputPath);
                }

                // Load workbook, add YoY growth column, and save result
                Workbook wb = new Workbook(inputPath);
                YoYGrowthCalculator.AddYoYGrowth(wb);
                wb.Save(outputPath);

                Console.WriteLine($"YoY Growth column added. Output saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
