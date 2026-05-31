using System;
using Aspose.Cells;

namespace AsposeCellsQuotePrefixValidation
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Access cell A1 and put a value that looks like a number
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("12345");

            // Create a style with QuotePrefix enabled and an initial number format
            Style style = workbook.CreateStyle();
            style.QuotePrefix = true;               // Enable quote prefix
            style.Custom = "0";                     // Simple number format

            // Create a StyleFlag that applies both QuotePrefix and NumberFormat
            StyleFlag flag = new StyleFlag();
            flag.QuotePrefix = true;
            flag.NumberFormat = true;

            // Apply the style to the cell
            cell.SetStyle(style, flag);

            // Verify that QuotePrefix is set
            Console.WriteLine("After first style application:");
            Console.WriteLine("QuotePrefix: " + cell.GetStyle().QuotePrefix);
            Console.WriteLine("Number Format: " + cell.GetStyle().Custom);

            // Change only the number format while keeping QuotePrefix flag true
            style.Custom = "#,##0";                 // New number format
            // Flag already has QuotePrefix = true; keep it unchanged
            flag.NumberFormat = true;               // Ensure number format flag is still true

            // Re-apply the updated style
            cell.SetStyle(style, flag);

            // Verify that QuotePrefix is still retained after the update
            Console.WriteLine("\nAfter updating number format:");
            Console.WriteLine("QuotePrefix: " + cell.GetStyle().QuotePrefix);
            Console.WriteLine("Number Format: " + cell.GetStyle().Custom);

            // Save the workbook (the file can be inspected manually if needed)
            workbook.Save("QuotePrefixValidation.xlsx");
        }
    }
}