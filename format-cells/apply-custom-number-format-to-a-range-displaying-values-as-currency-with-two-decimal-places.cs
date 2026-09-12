// Title: Apply a custom currency format ($#,##0.00) with two decimal places to a range of cells using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a workbook, defines a style with the custom number format "$#,##0.00", and applies it to the range A1:A5 using Aspose.Cells. | Write a .NET snippet that formats column A as currency with two decimal places by applying a style via Range.ApplyStyle together with a StyleFlag that enables NumberFormat.
// Common Searches: Aspose.Cells C# how to format a column as currency with two decimal places | set custom number format $#,##0.00 for multiple cells using Aspose.Cells .NET | apply currency style to a range in Excel with Aspose.Cells API | C# Aspose.Cells example for formatting cells as US dollars | range.ApplyStyle number format flag usage Aspose.Cells
// Tags: custom currency number format Aspose.Cells .NET | apply style to cell range Aspose.Cells | Range.ApplyStyle number format flag | Excel currency formatting C# Aspose | two decimal places number format Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a new workbook, inserts numeric values into cells A1‑A5, defines a style with the custom currency format "$#,##0.00", applies this style to the range A1:A5 using a StyleFlag that enables number formatting, and saves the file as CurrencyFormatted.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample numeric values in column A
            sheet.Cells["A1"].PutValue(1234.5);
            sheet.Cells["A2"].PutValue(5678.9);
            sheet.Cells["A3"].PutValue(0);
            sheet.Cells["A4"].PutValue(-2500.75);
            sheet.Cells["A5"].PutValue(99999.99);

            // Create a style with a custom currency number format
            Style currencyStyle = workbook.CreateStyle();
            currencyStyle.Custom = "$#,##0.00";

            // Apply the style to the range A1:A5
            Aspose.Cells.Range range = sheet.Cells.CreateRange("A1:A5");
            range.ApplyStyle(currencyStyle, new StyleFlag() { NumberFormat = true });

            // Save the workbook
            string outputPath = "CurrencyFormatted.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
