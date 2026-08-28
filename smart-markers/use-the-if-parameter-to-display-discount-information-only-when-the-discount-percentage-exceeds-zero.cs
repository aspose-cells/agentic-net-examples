// Title: Apply an IF formula with Aspose.Cells in C# to show discount amount only when the discount percentage is greater than zero
// AI Prompts: Write C# code using Aspose.Cells that inserts an IF formula to compute the discount amount only if the discount percentage cell is greater than zero, otherwise leaving the cell empty. | Create an Aspose.Cells workbook, format column B as a percentage and column C as currency, and add a conditional formula that returns an empty string when the discount percent is zero or negative.
// Common Searches: how to use Aspose.Cells C# IF function to hide zero discount values | Aspose.Cells C# calculate discount amount with conditional formula | format percentage column and currency column in Aspose.Cells workbook | set empty string result for IF formula in Aspose.Cells generated Excel | apply conditional discount calculation in Excel using Aspose.Cells C#
// Tags: Aspose.Cells IF formula conditional calculation | C# Aspose.Cells percentage column formatting | C# Aspose.Cells currency cell styling | Aspose.Cells generate Excel with discount logic | Aspose.Cells conditional display of values

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example creates a new workbook, adds headers for original price, discount %, and discount amount, inserts sample data, applies an IF formula in cell C2 that calculates A2*B2 only when B2>0 (otherwise returns an empty string), formats column B as a percentage and column C as currency, recalculates formulas, and saves the file as DiscountIfDemo.xlsx.
    public class DiscountIfDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Header row
                sheet.Cells["A1"].PutValue("Original Price");
                sheet.Cells["B1"].PutValue("Discount %");
                sheet.Cells["C1"].PutValue("Discount Amount");

                // Sample data
                sheet.Cells["A2"].PutValue(100);      // $100
                sheet.Cells["B2"].PutValue(0.15);     // 15% discount

                // IF formula: calculate discount amount only when discount > 0
                sheet.Cells["C2"].Formula = "=IF(B2>0, A2*B2, \"\")";

                // Format column B as percentage
                Style percentStyle = workbook.CreateStyle();
                percentStyle.Number = 10; // Built‑in percentage format
                sheet.Cells["B2"].SetStyle(percentStyle);

                // Format column C as currency
                Style currencyStyle = workbook.CreateStyle();
                currencyStyle.Number = 164; // Built‑in currency format with two decimals
                sheet.Cells["C2"].SetStyle(currencyStyle);

                // Recalculate formulas
                workbook.CalculateFormula();

                // Save the workbook
                string outputPath = "DiscountIfDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during DiscountIfDemo execution: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                DiscountIfDemo.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
