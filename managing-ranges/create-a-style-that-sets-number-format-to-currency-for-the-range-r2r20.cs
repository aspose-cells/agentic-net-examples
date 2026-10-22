// Title: How to apply a built‑in currency number format (ID 164) to cells R2:R20 using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set currency format for range R2:R20 | Apply built‑in number format ID 164 to a cell range with Aspose.Cells | Create and apply a currency style to Excel cells using Aspose.Cells .NET | Formatting a column as currency in Aspose.Cells workbook C# example | How to use StyleFlag to apply number format to a range in Aspose.Cells
// Tags: currency number format style Aspose.Cells | apply style to Excel range C# | built‑in number format ID 164 Aspose.Cells | StyleFlag NumberFormat property usage | format column as currency Aspose.Cells .NET

using Aspose.Cells;
using System;
using System.IO;

// The example creates a workbook, defines a style with the built‑in currency number format (ID 164), enables the NumberFormat flag, applies the style to cells R2:R20, and saves the workbook as Output.xlsx.
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

            // Create a new style object
            Style currencyStyle = workbook.CreateStyle();

            // Set the number format to a built‑in currency format (ID 164 = "$#,##0.00")
            currencyStyle.Number = 164;

            // Enable the number format flag so it will be applied
            StyleFlag flag = new StyleFlag
            {
                NumberFormat = true
            };

            // Define the target range R2:R20
            Aspose.Cells.Range targetRange = sheet.Cells.CreateRange("R2:R20");

            // Apply the currency style to the range
            targetRange.ApplyStyle(currencyStyle, flag);

            // Define output file path
            string outputPath = "Output.xlsx";

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
