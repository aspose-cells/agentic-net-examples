// Title: Apply a custom currency number format with two decimal places to a cell range using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# format range as currency with two decimal places | How to set custom number format "$#,##0.00" in Aspose.Cells | Applying a number format style to multiple cells with Aspose.Cells .NET | Create and apply a currency style to cells in Aspose.Cells C# example | Save workbook with currency‑formatted cells using Aspose.Cells
// Tags: currency number format Aspose.Cells | apply style to cell range Aspose.Cells | StyleFlag NumberFormat Aspose.Cells | CreateStyle custom format Aspose.Cells | Excel currency formatting .NET

using Aspose.Cells;
using System;

// The sample creates a workbook, writes numeric values to cells A1‑A3, defines a custom style with the "$#,##0.00" format, applies the style to the range A1:A3 using a StyleFlag that enables only number formatting, and saves the file as CurrencyFormatted.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Insert sample numeric values
        sheet.Cells["A1"].PutValue(1234.5);
        sheet.Cells["A2"].PutValue(9876.543);
        sheet.Cells["A3"].PutValue(-45.6);

        // Create a style for currency formatting with two decimal places
        Style currencyStyle = workbook.CreateStyle();
        // Custom number format: currency symbol with two decimal places
        currencyStyle.Custom = "$#,##0.00";

        // Define which style attributes to apply (only number format)
        StyleFlag flag = new StyleFlag();
        flag.NumberFormat = true;

        // Apply the currency style to the range containing the numbers
        sheet.Cells.CreateRange("A1:A3").ApplyStyle(currencyStyle, flag);

        // Save the workbook to a file
        workbook.Save("CurrencyFormatted.xlsx");
    }
}
