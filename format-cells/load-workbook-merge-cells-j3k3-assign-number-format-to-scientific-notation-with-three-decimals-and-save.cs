// Title: How to merge cells J3:K3 and apply a three‑decimal scientific number format using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# merge a range and set custom scientific format 0.000E+00 | apply three‑decimal scientific notation to merged cells in an Excel workbook with Aspose.Cells | C# code to merge J3 and K3 and format the range as scientific notation using Aspose.Cells | load workbook, merge cells, and save after applying custom number format in Aspose.Cells | set custom number format for a merged range in Excel via Aspose.Cells for .NET
// Tags: merge cells range Aspose.Cells C# | custom scientific number format Aspose.Cells | apply style to merged cells Aspose.Cells | Excel workbook formatting Aspose.Cells | save workbook after cell merge Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing Excel workbook, merges the cells J3:K3 on the first worksheet, creates a style with the custom scientific format "0.000E+00", applies this style to the merged range, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Merge cells J3:K3 (zero‑based indices: row 2, column 9)
            sheet.Cells.Merge(2, 9, 1, 2);

            // Create a style with scientific notation (three decimal places)
            Style sciStyle = workbook.CreateStyle();
            sciStyle.Custom = "0.000E+00";

            // Apply the style to the merged range
            Aspose.Cells.Range mergedRange = sheet.Cells.CreateRange("J3:K3");
            StyleFlag flag = new StyleFlag { NumberFormat = true };
            mergedRange.ApplyStyle(sciStyle, flag);

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
