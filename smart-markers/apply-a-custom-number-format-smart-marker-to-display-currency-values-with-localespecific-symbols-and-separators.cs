// Title: Apply a locale‑specific custom currency format to a cell range using Aspose.Cells smart markers in C#
// AI Prompts: Write C# code that sets Workbook.Settings.CultureInfo to a target locale, creates a Style with a custom currency pattern, and applies only the NumberFormat flag to cells A1:A3 via Aspose.Cells. | Show how to format negative monetary values in red while automatically using the locale's currency symbol in an Aspose.Cells workbook. | Demonstrate applying a custom number‑format style to a range through a smart‑marker without altering other cell attributes in Aspose.Cells.
// Common Searches: Aspose.Cells C# set workbook culture to French for currency formatting | How to apply a custom number format to a range of cells with Aspose.Cells | Display negative amounts in red using Aspose.Cells style flags | Use smart markers to format currency values based on locale in Aspose.Cells
// Tags: set workbook cultureinfo Aspose.Cells | custom currency number format style Aspose.Cells | apply numberformat flag to range Aspose.Cells | negative amount red formatting Aspose.Cells | smart marker locale currency Aspose.Cells

using System;
using System.Globalization;
using Aspose.Cells;

// Shows how to set the workbook culture, create a custom currency number format style, apply only the number‑format flag to cells A1:A3, and save the file using Aspose.Cells smart markers.
class CustomCurrencySmartMarkerDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Set the workbook culture to French (France) – this changes the currency symbol and separators
            wb.Settings.CultureInfo = new CultureInfo("fr-FR");

            // Access the first worksheet
            Worksheet ws = wb.Worksheets[0];

            // Populate some numeric values
            ws.Cells["A1"].PutValue(1234.56);
            ws.Cells["A2"].PutValue(7890.12);
            ws.Cells["A3"].PutValue(-345.67);

            // Create a style with a custom currency format
            // The "$" placeholder will be replaced by the locale‑specific symbol via CultureInfo
            Style style = wb.CreateStyle();
            style.Custom = "$#,##0.00;[Red]($#,##0.00)";

            // Apply only the number format part of the style
            StyleFlag flag = new StyleFlag();
            flag.NumberFormat = true;

            // Apply the style to the range A1:A3
            Aspose.Cells.Range range = ws.Cells.CreateRange("A1:A3");
            range.ApplyStyle(style, flag);

            // Save the workbook
            wb.Save("CustomCurrencySmartMarker.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
