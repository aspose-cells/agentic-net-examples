// Title: C# routine to delete all WordArt shapes and add a diagonal CONFIDENTIAL text‑effect watermark to each worksheet using Aspose.Cells
// AI Prompts: Generate C# code that iterates through every worksheet, removes any shape whose IsWordArt property is true, and then creates a rotated gray text‑effect shape with the text "CONFIDENTIAL" that spans the whole sheet using Aspose.Cells. | Provide a method that clears existing WordArt from an Excel workbook and inserts a diagonal watermark by calling Shape.AddTextEffect, setting rotation, size, and transparency with Aspose.Cells for .NET.
// Common Searches: aspnet remove wordart shapes from excel worksheet using aspose.cells | how to add diagonal text watermark to all sheets in an excel file with aspose.cells c# | c# delete existing wordart before inserting new watermark in aspose.cells workbook | configure text effect shape size to cover entire worksheet aspose.cells | aspose.cells set shape transparency for watermark in c#
// Tags: remove WordArt shapes Aspose.Cells C# | add rotated text effect watermark Aspose.Cells | shape removal before watermark insertion Aspose.Cells | adjust text effect dimensions to cover worksheet Aspose.Cells | set shape transparency Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsWatermark
{
    // The program loads an Excel workbook, removes every WordArt shape from each worksheet, adds a diagonal gray "CONFIDENTIAL" text‑effect shape that covers the full sheet, and saves the updated file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Process each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    try
                    {
                        // -------------------------------------------------
                        // Remove all existing WordArt shapes from the sheet
                        // -------------------------------------------------
                        for (int i = sheet.Shapes.Count - 1; i >= 0; i--)
                        {
                            Shape shape = sheet.Shapes[i];
                            // Use IsWordArt property to identify WordArt shapes
                            if (shape.IsWordArt)
                            {
                                sheet.Shapes.RemoveAt(i);
                            }
                        }

                        // -------------------------------------------------
                        // Add a new watermark as TextEffect (fallback for WordArt)
                        // -------------------------------------------------
                        // Create a text effect shape with the desired text
                        Shape watermark = sheet.Shapes.AddTextEffect(
                            MsoPresetTextEffect.TextEffect1, // preset effect
                            "CONFIDENTIAL",                  // text
                            "Arial",                         // font name
                            72,                              // font size
                            false,                           // bold
                            false,                           // italic
                            0,                               // left (pixels)
                            0,                               // top (pixels)
                            0,                               // width (will be adjusted later)
                            0,                               // height (will be adjusted later)
                            0,                               // anchor row
                            0);                              // anchor column

                        // Configure visual appearance of the watermark
                        watermark.RotationAngle = -45;                     // diagonal orientation
                        watermark.Font.Color = Color.Gray;                // gray color
                        watermark.Font.Size = 72;                         // large font size
                        watermark.Fill.Transparency = 0.5;                // 50% transparent

                        // Position the watermark to cover the sheet
                        int maxColumn = sheet.Cells.MaxColumn + 1;
                        int maxRow = sheet.Cells.MaxRow + 1;
                        double columnWidth = sheet.Cells.StandardWidth * 256; // points per column
                        double rowHeight = sheet.Cells.StandardHeight * 15;   // points per row

                        watermark.Width = (int)(maxColumn * columnWidth);
                        watermark.Height = (int)(maxRow * rowHeight);
                        watermark.Top = 0;
                        watermark.Left = 0;
                    }
                    catch (Exception exSheet)
                    {
                        Console.WriteLine($"Error processing sheet '{sheet.Name}': {exSheet.Message}");
                    }
                }

                // Save the modified workbook
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved successfully to {outputPath}");
                }
                catch (Exception exSave)
                {
                    Console.WriteLine($"Failed to save workbook: {exSave.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
