// Title: Copy a built‑in currency number format from a source cell to a target range using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells in C# to copy the currency number format from cell A1 to the range B1:D3 while preserving the $ symbol. | Apply a built‑in number style (Number = 164) to multiple cells in an Excel workbook with Aspose.Cells for .NET. | Create a workbook, set a currency format on a source cell, and propagate that style to a specified range using the Aspose.Cells C# API.
// Common Searches: how to copy a currency number format to a range with Aspose.Cells C# | preserve built‑in Excel number formatting when applying style to multiple cells in .NET | Aspose.Cells copy style from one cell to another range example | set $ currency format for a block of cells using Aspose.Cells for .NET
// Tags: currency number format copying Aspose.Cells | built‑in number style application C# | cell range number format assignment .NET | Excel currency formatting preservation Aspose.Cells | style inheritance for multiple cells Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new workbook, applies the built‑in currency number format (Number = 164) to cell A1, defines the target range B1:D3, copies the source style to each cell in that range, and saves the workbook as CopyStyleCurrency.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle create)
            var workbook = new Workbook();

            // Access the first worksheet
            var sheet = workbook.Worksheets[0];

            // Set up a source cell with a currency number format
            var sourceCell = sheet.Cells["A1"];
            sourceCell.PutValue(1234.56);
            var sourceStyle = sourceCell.GetStyle();
            // Use built‑in currency format (e.g., $#,##0.00)
            sourceStyle.Number = 164;
            sourceCell.SetStyle(sourceStyle);

            // Define the target range where the style will be copied
            var targetRange = sheet.Cells.CreateRange("B1:D3");

            // Apply the source style to each cell in the target range
            foreach (Cell cell in targetRange)
            {
                cell.SetStyle(sourceStyle);
            }

            // Ensure the output directory exists
            string outputPath = "CopyStyleCurrency.xlsx";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (lifecycle save)
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
