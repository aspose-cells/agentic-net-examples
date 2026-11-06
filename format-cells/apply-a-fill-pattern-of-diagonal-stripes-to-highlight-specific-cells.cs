// Title: How to apply a diagonal stripe fill pattern to a range of cells using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set diagonal stripe background for a specific cell range | Apply diagonal stripe fill style to Excel cells with Aspose.Cells .NET | Highlight cells B3:D4 with diagonal stripe pattern using Aspose.Cells in C# | Use StyleFlag CellShading to add diagonal stripe shading in Aspose.Cells workbook | Create custom cell style with diagonal stripe pattern in Aspose.Cells for .NET
// Tags: diagonal stripe fill pattern Aspose.Cells | apply style to range B3:D4 | StyleFlag CellShading property | custom cell shading style .NET | highlight Excel cells using Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a workbook, fills cells B2:D4 with data, defines a style with a LightBlue foreground, White background, and a diagonal stripe pattern, applies this style to the range B3:D4 using a StyleFlag that enables cell shading, and saves the file as HighlightedCells.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Fill cells B2:D4 with sample data
                sheet.Cells["B2"].PutValue("Item 1");
                sheet.Cells["C2"].PutValue("Item 2");
                sheet.Cells["D2"].PutValue("Item 3");
                sheet.Cells["B3"].PutValue(10);
                sheet.Cells["C3"].PutValue(20);
                sheet.Cells["D3"].PutValue(30);
                sheet.Cells["B4"].PutValue(40);
                sheet.Cells["C4"].PutValue(50);
                sheet.Cells["D4"].PutValue(60);

                // Define a style with diagonal stripe fill pattern
                Style diagonalStyle = workbook.CreateStyle();
                diagonalStyle.ForegroundColor = Color.LightBlue;   // Stripe color
                diagonalStyle.BackgroundColor = Color.White;      // Base color
                diagonalStyle.Pattern = BackgroundType.DiagonalStripe; // Diagonal stripe pattern

                // Apply the style to the target range B3:D4
                Aspose.Cells.Range targetRange = sheet.Cells.CreateRange("B3:D4");
                StyleFlag flag = new StyleFlag { CellShading = true };
                targetRange.ApplyStyle(diagonalStyle, flag);

                // Save the workbook
                string outputPath = "HighlightedCells.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
