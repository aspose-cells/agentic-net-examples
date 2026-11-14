// Title: Copy currency number format from one cell to multiple cells with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# copy custom currency format from a source cell to a range | How to preserve $ symbol when copying cell styles in Aspose.Cells .NET | SetStyle GetStyle example for duplicating number formats in Excel using Aspose.Cells | Copy style with currency formatting to B1:B5 in a generated workbook Aspose.Cells
// Tags: copy cell style Aspose.Cells C# | currency number format SetStyle | duplicate Excel cell formatting Aspose.Cells | preserve custom number format .NET | apply style to range Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCopyStyleExample
{
    // The example creates a workbook, defines a custom currency format ("$#,##0.00") on cell A1, then copies that style to cells B1 through B5 using GetStyle() and SetStyle(), ensuring the $ symbol and formatting are retained, and saves the file as CopyStyleCurrencyExample.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Define a style with a currency number format
                Style currencyStyle = workbook.CreateStyle();
                // Use a custom format string for currency
                currencyStyle.Custom = "$#,##0.00";

                // Apply the style to a source cell (e.g., A1) and set a value
                Cell sourceCell = sheet.Cells["A1"];
                sourceCell.PutValue(1234.56);
                sourceCell.SetStyle(currencyStyle);

                // Copy the style to target cells (B1:B5)
                for (int row = 0; row < 5; row++)
                {
                    Cell targetCell = sheet.Cells[row, 1]; // Column B (index 1)
                    targetCell.PutValue(1234.56 + row); // Different values to show formatting

                    // Copy the style from the source cell to the target cell
                    // Aspose.Cells does not have Cell.CopyStyle(sourceCell); use SetStyle with the source style
                    targetCell.SetStyle(sourceCell.GetStyle());
                }

                // Define output file path
                string outputPath = "CopyStyleCurrencyExample.xlsx";

                // Save the workbook (lifecycle rule: save)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
