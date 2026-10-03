// Title: How to use Aspose.Cells for .NET to conditionally display discount text with an IF formula when the percentage is greater than zero
// AI Prompts: Generate C# code that creates an Excel workbook, writes discount values as percentages in column B, and inserts an IF formula in column C that outputs "Discount: xx%" only for rows where the percentage exceeds zero, using Aspose.Cells. | Write a C# example that applies a percentage number format to cells, builds a row‑specific IF formula referencing each B‑cell, and saves the file as DiscountInfo.xlsx with Aspose.Cells.
// Common Searches: Aspose.Cells C# conditional IF formula to show text only for non‑zero percentages | Create Excel file with discount column and conditional discount label using Aspose.Cells | How to format cells as percentages and add IF text in Aspose.Cells .NET | Saving workbook with conditional discount messages in Aspose.Cells | Using CellsHelper.CellIndexToName with IF formula in Aspose.Cells C#
// Tags: Aspose.Cells conditional IF formula | percentage number format Aspose.Cells | write discount percentages Aspose.Cells | dynamic cell address formula Aspose.Cells | save workbook as .xlsx Aspose.Cells | C# Aspose.Cells smart markers example

using Aspose.Cells;
using System;
using System.IO;

// The sample creates a new workbook, writes discount percentages as fractional values formatted as percentages in column B, adds an IF formula in column C that displays "Discount: xx%" only when the corresponding percentage is greater than zero, and saves the result as DiscountInfo.xlsx.
class DiscountDisplay
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Example discount percentages (some zero, some non‑zero)
            double[] discounts = { 0.0, 5.5, 0.0, 12.75, 0.0 };

            // Prepare a style for percentage formatting (Number format index 10)
            Style percentStyle = workbook.CreateStyle();
            percentStyle.Number = 10; // Percentage format

            // Populate column B with discount percentages (as fractions) and apply style
            for (int i = 0; i < discounts.Length; i++)
            {
                // Store as fraction (e.g., 0.055 for 5.5%)
                sheet.Cells[i, 1].PutValue(discounts[i] / 100);
                sheet.Cells[i, 1].SetStyle(percentStyle);
            }

            // Column C will display discount information only when the percentage > 0
            for (int i = 0; i < discounts.Length; i++)
            {
                // Build the IF formula: if B cell > 0 then show text, else empty string
                string bAddress = CellsHelper.CellIndexToName(i, 1); // e.g., B1, B2...
                string formula = $"IF({bAddress}>0, \"Discount: \" & TEXT({bAddress},\"0.00%\"), \"\")";
                sheet.Cells[i, 2].Formula = formula;
            }

            // Define output file path
            string outputPath = "DiscountInfo.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred while creating the workbook:");
            Console.WriteLine(ex.Message);
        }
    }
}
