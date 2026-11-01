// Title: How to format a range of cells as currency with two decimal places using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# apply $#,##0.00 currency format to multiple cells | set custom number format with two decimal places for a range in Aspose.Cells | C# Aspose.Cells style flag all property to format Excel cells as currency | apply currency number format to A1:A3 using Aspose.Cells .NET API | format numeric values as currency in Aspose.Cells workbook
// Tags: custom currency number format Aspose.Cells | range style application C# | StyleFlag All attribute Aspose.Cells | Excel $#,##0.00 formatting .NET | Aspose.Cells workbook currency styling

using System;
using Aspose.Cells;

// Creates a workbook, inserts numeric values into A1‑A3, defines a custom style with the "$#,##0.00" currency format, applies the style to the range using a StyleFlag with All set to true, and saves the file as FormattedCurrency.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Get the first worksheet
            var sheet = workbook.Worksheets[0];

            // Sample numeric values
            sheet.Cells["A1"].PutValue(1234.567);
            sheet.Cells["A2"].PutValue(89.1);
            sheet.Cells["A3"].PutValue(-45.678);

            // Define a custom number format for currency with two decimal places ("$#,##0.00")
            var style = workbook.CreateStyle();
            style.Custom = "$#,##0.00";

            // Apply the style to the target range
            var styleFlag = new StyleFlag();
            styleFlag.All = true; // Apply all style attributes, including the custom number format

            sheet.Cells.CreateRange("A1:A3").ApplyStyle(style, styleFlag);

            // Save the workbook
            workbook.Save("FormattedCurrency.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
