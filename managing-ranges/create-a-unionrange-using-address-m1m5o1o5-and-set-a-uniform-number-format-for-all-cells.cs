// Title: Apply a Uniform Number Format to a UnionRange (M1:M5, O1:O5) with Aspose.Cells for .NET
// Description: Shows how to create a UnionRange that includes M1:M5 and O1:O5 on a worksheet, define a style with the custom numeric format "0.00", apply the style to the whole range using StyleFlag, and save the workbook as UnionRangeNumberFormat.xlsx.
// Keywords: Aspose.Cells | C# UnionRange | non‑contiguous range | custom numeric format | 0.00 format | .NET spreadsheet | apply style to multiple ranges | StyleFlag | Excel cell formatting | UnionRangeNumberFormat.xlsx
// Common Searches: Aspose.Cells create union range C# | set number format for disjoint cells Aspose | apply style to multiple ranges .NET | custom numeric format 0.00 Aspose.Cells | save workbook after formatting union range
// Developer Intent: Create a UnionRange covering M1:M5 and O1:O5 and assign the same numeric format to every cell in the range.
// Use Cases: Standardize financial figures in two separate columns before generating a report. | Ensure consistent decimal precision across non‑adjacent data sections in an exported spreadsheet. | Prepare a worksheet for conditional formatting or charting by applying a uniform number format to multiple disjoint ranges.
// AI Prompts: Write C# code using Aspose.Cells to create a UnionRange from M1:M5 and O1:O5 and set the number format to "0.00" for all cells. | Explain how StyleFlag determines which style attributes are applied when formatting a UnionRange in Aspose.Cells. | Provide an example that formats a UnionRange with a custom numeric pattern, saves the workbook, and verifies the applied format.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsUnionRangeDemo
{
    // Shows how to create a UnionRange that includes M1:M5 and O1:O5 on a worksheet, define a style with the custom numeric format "0.00", apply the style to the whole range using StyleFlag, and save the workbook as UnionRangeNumberFormat.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Create a union range that includes columns M and O rows 1 to 5
                UnionRange unionRange = workbook.Worksheets.CreateUnionRange("M1:M5,O1:O5", 0);

                // Define a style with a uniform number format (e.g., two decimal places)
                Style numberStyle = workbook.CreateStyle();
                numberStyle.Custom = "0.00";

                // Apply only the number format to the union range
                StyleFlag flag = new StyleFlag();
                flag.All = true; // Applying all style attributes; only number format is set

                unionRange.ApplyStyle(numberStyle, flag);

                // Save the workbook
                string outputPath = "UnionRangeNumberFormat.xlsx";
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
