// Title: How to duplicate an Excel worksheet with all shapes and keep their Z‑order using Aspose.Cells for .NET
// AI Prompts: Use Aspose.Cells Worksheet.Copy to create a copy of 'Sheet1' named 'Sheet1_Copy' while retaining the original Z‑order of every shape. | Programmatically clone a worksheet in C# so that cells, formats, and shape layering are identical to the source sheet. | Generate a new worksheet that mirrors an existing one and preserves drawing order of embedded shapes when saving the workbook with Aspose.Cells.
// Common Searches: Aspose.Cells copy worksheet with shapes preserving Z order C# | duplicate Excel sheet and keep shape layering using .NET library | clone worksheet including drawings and maintain Z‑order Aspose.Cells | Worksheet.Copy retains shape order example Aspose.Cells
// Tags: Worksheet.Copy preserving shape Z-order | duplicate worksheet with drawings Aspose.Cells | clone Excel sheet maintaining shape layering C# | copy sheet including shapes Aspose.Cells .NET | preserve drawing order when copying worksheet

using Aspose.Cells;
using System;
using System.IO;

// The example loads source.xlsx, uses Worksheet.Copy to duplicate the 'Sheet1' worksheet into a new sheet named 'Sheet1_Copy', preserving cells, formats, and all shapes with their original Z‑order, and then saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            string sourcePath = "source.xlsx";
            string outputPath = "output.xlsx";

            // Ensure the source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(sourcePath);

            // Get the worksheet to duplicate
            Worksheet srcSheet = workbook.Worksheets["Sheet1"];
            if (srcSheet == null)
            {
                Console.WriteLine("Worksheet 'Sheet1' not found.");
                return;
            }

            // Add a new worksheet for the copy
            Worksheet destSheet = workbook.Worksheets.Add("Sheet1_Copy");

            // Copy the entire sheet (cells, formats, shapes, etc.) preserving Z‑order
            srcSheet.Copy(destSheet);

            // Save the workbook with the duplicated worksheet
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
