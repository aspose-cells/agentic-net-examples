// Title: Set a custom RGB fill color for header cells in Aspose.Cells (C#)
// Description: Creates a workbook, defines the A1:D1 header range, builds a SteelBlue color with RGB (70,130,180), applies a solid fill style using StyleFlag.CellShading, writes header text, and saves the file as HeaderFillColorDemo.xlsx.
// Keywords: Aspose.Cells C# fill color | custom RGB background Aspose.Cells | header row style Aspose.Cells | StyleFlag cell shading | solid fill pattern Excel .NET | create CellsColor from RGB
// Common Searches: how to set custom RGB fill color in Aspose.Cells C# | apply solid background to a range with Aspose.Cells | use StyleFlag to enable cell shading in Aspose.Cells | create reusable header style Aspose.Cells .NET | Aspose.Cells set header background color programmatically
// Developer Intent: Apply a specific RGB color as a solid fill to a header row in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Brand‑consistent Excel reports with corporate RGB header colors. | Visually distinct table headers in generated spreadsheets. | Reusable header styling across multiple worksheets in a single workbook.
// AI Prompts: Generate C# code that uses Aspose.Cells to apply a custom RGB fill (70,130,180) with a solid pattern to cells A1:D1. | Show how to create a reusable Style containing a custom fill color and apply it to header rows in several worksheets. | Explain converting a hex color (e.g., #4682B4) to System.Drawing.Color and using it as a CellsColor for header styling in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demo class that creates a workbook with a colored header row
    // Creates a workbook, defines the A1:D1 header range, builds a SteelBlue color with RGB (70,130,180), applies a solid fill style using StyleFlag.CellShading, writes header text, and saves the file as HeaderFillColorDemo.xlsx.
    public class HeaderFillColorDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Define the header range (cells A1 to D1)
                Aspose.Cells.Range headerRange = worksheet.Cells.CreateRange("A1", "D1");

                // Create a custom RGB color (SteelBlue)
                CellsColor headerColor = workbook.CreateCellsColor();
                headerColor.Color = Color.FromArgb(70, 130, 180);

                // Create a style and apply the custom fill color
                Style headerStyle = workbook.CreateStyle();
                headerStyle.ForegroundColor = headerColor.Color;
                headerStyle.Pattern = BackgroundType.Solid; // Solid fill pattern

                // Apply the style to the header range (cell shading flag)
                StyleFlag flag = new StyleFlag { CellShading = true };
                headerRange.ApplyStyle(headerStyle, flag);

                // Add header text
                worksheet.Cells["A1"].PutValue("Header 1");
                worksheet.Cells["B1"].PutValue("Header 2");
                worksheet.Cells["C1"].PutValue("Header 3");
                worksheet.Cells["D1"].PutValue("Header 4");

                // Save the workbook
                string outputPath = "HeaderFillColorDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    internal class Program
    {
        private static void Main(string[] args)
        {
            HeaderFillColorDemo.Run();
        }
    }
}
