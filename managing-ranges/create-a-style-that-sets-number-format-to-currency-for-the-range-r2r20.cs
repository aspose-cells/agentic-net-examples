// Title: How to apply a currency number format style to the range R2:R20 with Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a Style with a custom currency format ($#,##0.00) and applies it to the range R2:R20 using Aspose.Cells. | Show how to configure a StyleFlag to apply all style attributes and use ApplyStyle to set the number format for a specific range. | Demonstrate saving the workbook after the currency style has been applied to the selected cells.
// Common Searches: Aspose.Cells C# set currency format for cells R2 through R20 | apply custom number format to a column range using Aspose.Cells .NET | StyleFlag all attributes example Aspose.Cells C# | how to format Excel cells as currency with Aspose.Cells programmatically | C# Aspose.Cells apply style to range R2:R20 and save workbook
// Tags: currency number format Aspose.Cells | apply style to range Aspose.Cells | StyleFlag apply all attributes | custom Excel number format C# | set number format column R Aspose.Cells

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example creates a new workbook, defines a Style with the custom currency format "$#,##0.00", builds the range R2:R20, applies the style to the range using a StyleFlag that enables all attributes, and saves the workbook as StyledWorkbook.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a style with a custom currency format
            Style currencyStyle = workbook.CreateStyle();
            currencyStyle.Custom = "$#,##0.00";

            // Define the range R2:R20 (use alias to avoid conflict with System.Range)
            AsposeRange range = worksheet.Cells.CreateRange("R2", "R20");

            // Specify which style attributes to apply (apply all to include number format)
            StyleFlag flag = new StyleFlag();
            flag.All = true;

            // Apply the style to the range
            range.ApplyStyle(currencyStyle, flag);

            // Save the workbook
            workbook.Save("StyledWorkbook.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
