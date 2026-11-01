// Title: How to format a column as currency and freeze the totals rows in an Aspose.Cells workbook using C#
// AI Prompts: Write C# code that applies the built‑in currency number format (ID 164) to a given cell range and then freezes the first N rows in an Aspose.Cells workbook. | Show how to set a monetary style for cells B2:B4 and keep the totals row visible by using the FreezePanes method in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# apply built‑in currency format to a range of cells | C# freeze top rows in Excel workbook with Aspose.Cells FreezePanes | How to format monetary values and freeze header rows using Aspose.Cells .NET | Set number format 164 for column B and freeze rows 1‑4 in Aspose.Cells | Aspose.Cells example for currency formatting and row freeze in C#
// Tags: currency style number 164 Aspose.Cells | freeze first rows Aspose.Cells .NET | apply monetary style to cell range C# | Excel workbook row freeze after formatting | Aspose.Cells monetary formatting example

using System;
using Aspose.Cells;

// The program creates a new workbook, fills sample data, applies the built‑in currency format (Number 164) to cells B2:B4, freezes the first four rows so the totals row remains visible, and saves the file as MonetaryAndFreeze.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data with a totals row at row 4 (index 3)
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["A2"].PutValue("Product A");
            sheet.Cells["B2"].PutValue(1234.56);
            sheet.Cells["A3"].PutValue("Product B");
            sheet.Cells["B3"].PutValue(789.01);
            sheet.Cells["A4"].PutValue("Total");
            sheet.Cells["B4"].PutValue(1234.56 + 789.01);

            // Define a monetary number format (built‑in currency format, e.g., $#,##0.00)
            Style moneyStyle = workbook.CreateStyle();
            moneyStyle.Number = 164; // Built‑in currency format

            // Apply the monetary format to the amount column (B2:B4)
            StyleFlag flag = new StyleFlag { NumberFormat = true };
            sheet.Cells.CreateRange("B2:B4").ApplyStyle(moneyStyle, flag);

            // Freeze rows 1‑4 so the totals row stays visible while scrolling.
            // FreezePanes(row, column, totalRows, totalColumns)
            // row/column specify the first scrollable cell (row 5, column 0),
            // totalRows/totalColumns specify how many rows/columns to freeze.
            sheet.FreezePanes(5, 0, 4, 0);

            // Save the workbook
            workbook.Save("MonetaryAndFreeze.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
