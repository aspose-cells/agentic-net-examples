// Title: How to merge cells J3:K3 and apply a three‑decimal scientific notation format using Aspose.Cells for .NET
// AI Prompts: Load an existing Excel workbook, merge the cells spanning columns J and K in row 3 on the first worksheet, and assign a custom number format that displays values in scientific notation with three decimal places using Aspose.Cells. | Create a style with the custom format "0.000E+00", configure a StyleFlag to affect only NumberFormat, apply the style to the merged range, and save the workbook to a new file.
// Common Searches: Aspose.Cells C# merge J3 K3 and set scientific notation format | apply custom number format 0.000E+00 to merged cells using Aspose.Cells .NET | how to use StyleFlag to change only number format for a merged range in Aspose.Cells | save workbook after merging cells and applying scientific number format with Aspose.Cells
// Tags: merge cells range Aspose.Cells .NET | custom scientific number format Aspose.Cells | StyleFlag NumberFormat Aspose.Cells | apply style to merged range C# | save workbook after formatting Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// // Loads "input.xlsx", merges cells J3:K3 on the first worksheet, applies the custom scientific notation format "0.000E+00" to the merged range, and saves the result as "output.xlsx".
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (index 0)
            Worksheet sheet = workbook.Worksheets[0];

            // Merge cells J3:K3 (one additional column to the right of J3)
            Aspose.Cells.Range mergeRange = sheet.Cells.CreateRange("J3:K3");
            mergeRange.Merge();

            // Create a style that formats numbers in scientific notation with three decimals
            Style sciStyle = workbook.CreateStyle();
            sciStyle.Custom = "0.000E+00";

            // Apply the style to the merged range (only number format part)
            StyleFlag flag = new StyleFlag
            {
                NumberFormat = true
            };
            mergeRange.ApplyStyle(sciStyle, flag);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
