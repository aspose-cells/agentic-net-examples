// Title: How to duplicate an Excel range reference using Aspose.Cells for .NET by offsetting zero rows and columns
// Common Searches: Aspose.Cells C# duplicate range reference without creating a new range | offset a range by zero rows and columns using Aspose.Cells | apply yellow background to a copied range in Aspose.Cells .NET | reuse existing range object for styling in Aspose.Cells | create a second reference to the same cell range in C# Aspose.Cells
// Tags: duplicate range reference Aspose.Cells .NET | offset range zero rows columns Aspose.Cells | apply style to range Aspose.Cells | range copy without new object C# | Excel range styling Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads or creates an Excel workbook, defines the range A1:C3, creates a duplicate reference by assigning the original range to a new variable, applies a yellow background style to the duplicate range, and saves the result to output.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one with sample data
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                ws.Cells["A1"].PutValue("A1");
                ws.Cells["B1"].PutValue("B1");
                ws.Cells["C1"].PutValue("C1");
                ws.Cells["A2"].PutValue("A2");
                ws.Cells["B2"].PutValue("B2");
                ws.Cells["C2"].PutValue("C2");
                ws.Cells["A3"].PutValue("A3");
                ws.Cells["B3"].PutValue("B3");
                ws.Cells["C3"].PutValue("C3");
            }

            Worksheet sheet = workbook.Worksheets[0];

            // Create original range
            Aspose.Cells.Range originalRange = sheet.Cells.CreateRange("A1:C3");

            // Duplicate reference (same range)
            Aspose.Cells.Range duplicateRange = originalRange;

            // Apply yellow background to the duplicate range
            Style style = workbook.CreateStyle();
            style.ForegroundColor = Color.Yellow;
            style.Pattern = BackgroundType.Solid;

            StyleFlag flag = new StyleFlag
            {
                CellShading = true
            };

            duplicateRange.ApplyStyle(style, flag);

            // Ensure output directory exists
            string? outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
