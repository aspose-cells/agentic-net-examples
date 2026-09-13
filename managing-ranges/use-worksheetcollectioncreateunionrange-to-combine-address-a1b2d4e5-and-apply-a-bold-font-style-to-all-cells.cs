// Title: Create a union range for cells A1:B2 and D4:E5 and apply a bold font style with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses WorksheetCollection.CreateUnionRange to merge the ranges A1:B2 and D4:E5 and then applies a bold font to every cell in the union using Aspose.Cells. | Show how to define a Style and StyleFlag in Aspose.Cells and apply them to a non‑contiguous union range created from multiple addresses. | Explain the steps required to create a union range and style it bold in an in‑memory workbook with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# create union range from multiple addresses and set bold font | How to apply a style to non‑contiguous cells using CreateUnionRange in Aspose.Cells | C# example for formatting union range A1:B2,D4:E5 with Aspose.Cells | Apply bold text to a union range in an Aspose.Cells workbook | WorksheetCollection.CreateUnionRange usage for styling cells in .NET
// Tags: grouped cell range bold styling Aspose.Cells | C# discontiguous cells format Aspose.Cells | WorksheetCollection range grouping example | apply style to multiple address blocks Aspose.Cells | save XLSX workbook with styled cell groups

using Aspose.Cells;
using System;
using System.IO;

// // This program creates an in‑memory workbook, builds a union range covering A1:B2 and D4:E5 on the first worksheet, defines a bold font style, applies it to all cells in the union range, and saves the workbook to Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (in-memory)
            Workbook workbook = new Workbook();

            // Ensure at least one worksheet exists
            Worksheet sheet = workbook.Worksheets[0];

            // Create a union range that combines the two address blocks on the first sheet (index 0)
            UnionRange unionRange = workbook.Worksheets.CreateUnionRange("A1:B2,D4:E5", 0);

            // Define a style with a bold font
            Style boldStyle = workbook.CreateStyle();
            boldStyle.Font.IsBold = true;

            // Prepare a StyleFlag to apply all style attributes
            StyleFlag flag = new StyleFlag { All = true };

            // Apply the bold style to all cells in the union range
            unionRange.ApplyStyle(boldStyle, flag);

            // Define output file path
            string outputPath = "Output.xlsx";

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
