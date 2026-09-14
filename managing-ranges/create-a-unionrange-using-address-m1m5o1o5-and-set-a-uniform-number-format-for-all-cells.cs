// Title: Create a union range (M1:M5, O1:O5) and apply a two‑decimal number format to all cells with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells to define a named union range that includes M1:M5 and O1:O5 and apply a custom number format of "0.00" to every cell in the range. | Show how to use StyleFlag to apply a uniform style to a multi‑area range and then save the workbook as an .xlsx file.
// Common Searches: Aspose.Cells how to create a union range with non‑contiguous cells in C# | set the same custom number format for multiple separate ranges using Aspose.Cells .NET | apply StyleFlag.All to a named range that spans M1:M5 and O1:O5 in Aspose.Cells
// Tags: union range creation Aspose.Cells C# | apply custom number format Aspose.Cells | StyleFlag All property usage | save workbook as xlsx Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// // This program creates a new workbook, defines a union range covering M1:M5 and O1:O5, applies the custom number format "0.00" to all cells in that range, and saves the file as UnionRangeExample.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a union range that includes cells M1:M5 and O1:O5
            AsposeRange unionRange = worksheet.Cells.CreateRange("MyUnion", "M1:M5,O1:O5");

            // Define a uniform number format (two decimal places)
            Style uniformStyle = workbook.CreateStyle();
            uniformStyle.Custom = "0.00";

            // Apply the style to all cells in the union range
            StyleFlag styleFlag = new StyleFlag { All = true };
            unionRange.ApplyStyle(uniformStyle, styleFlag);

            // Define output file path
            string outputPath = "UnionRangeExample.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
