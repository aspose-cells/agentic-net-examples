// Title: How to format a range of cells as currency with two decimal places using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# format cells as currency with two decimal places | Apply custom number format $#,##0.00 to a range in Aspose.Cells | Use StyleFlag to change only number format in an Aspose.Cells workbook | Create Excel file with currency formatting using Aspose.Cells .NET | Set custom currency format for multiple cells in Aspose.Cells C#
// Tags: custom currency number format Aspose.Cells | apply number format to range Aspose.Cells C# | StyleFlag NumberFormat property Aspose.Cells | CreateRange A1:A5 formatting Aspose.Cells | save workbook as Excel file Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new workbook, writes numeric values to column A, defines the range A1:A5, builds a style with the custom currency format "$#,##0.00", applies only the number format via a StyleFlag, and saves the workbook as FormattedCurrency.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Sample numeric data to demonstrate the format
            double[] values = { 1234.5, 5678.9, 0, -250.75, 10000 };
            for (int i = 0; i < values.Length; i++)
            {
                // Place values in column A (cells A1, A2, ...)
                sheet.Cells[i, 0].PutValue(values[i]);
            }

            // Define the range that will display as currency
            Aspose.Cells.Range range = sheet.Cells.CreateRange("A1:A5");

            // Create a style object with custom currency format
            Style currencyStyle = workbook.CreateStyle();
            currencyStyle.Custom = "$#,##0.00";

            // Apply only the number format part of the style
            StyleFlag flag = new StyleFlag { NumberFormat = true };
            range.ApplyStyle(currencyStyle, flag);

            // Save the workbook to a file
            string outputFile = "FormattedCurrency.xlsx";
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputFile)}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
