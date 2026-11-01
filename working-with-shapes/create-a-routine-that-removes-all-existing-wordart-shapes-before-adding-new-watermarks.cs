// Title: How to remove existing WordArt shapes and add a centered semi‑transparent text watermark in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# remove all WordArt objects from a worksheet before adding a watermark | C# code to delete WordArt shapes in an Excel file using Aspose.Cells | Add a diagonal text watermark in Excel after clearing existing WordArt with Aspose.Cells | How to iterate the Shapes collection and identify WordArt in Aspose.Cells .NET | Center a WordArt watermark in the used range of an Excel sheet with Aspose.Cells
// Tags: delete WordArt objects Aspose.Cells | insert centered text watermark Aspose.Cells | identify WordArt via IsWordArt Aspose.Cells | compute worksheet used range center Aspose.Cells | apply shape transparency rotation Aspose.Cells

using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace Example
{
    // The example loads an Excel workbook, loops through each worksheet, gathers indices of shapes where the IsWordArt property is true, removes those shapes in reverse order, calculates the center of the used range, adds a new WordArt shape as a watermark with preset style, 50% fill transparency, -45° rotation, and no border, then saves the updated workbook.
    public class WatermarkHelper
    {
        /// <param name="inputPath">Full path to the source Excel file.</param>
        /// <param name="outputPath">Full path where the modified file will be saved.</param>
        /// <param name="watermarkText">Text to use for the new watermark.</param>
        public static void RemoveWordArtAndAddWatermark(string inputPath, string outputPath, string watermarkText)
        {
            try
            {
                if (!File.Exists(inputPath))
                    throw new FileNotFoundException($"Input file not found: {inputPath}");

                // Load the workbook (lifecycle rule: load)
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Collect indices of WordArt shapes to remove
                    List<int> indicesToRemove = new List<int>();

                    // Examine each shape on the worksheet
                    for (int i = 0; i < sheet.Shapes.Count; i++)
                    {
                        Shape shape = sheet.Shapes[i];

                        // Use IsWordArt property to identify WordArt objects
                        if (shape.IsWordArt)
                        {
                            indicesToRemove.Add(i);
                        }
                    }

                    // Remove shapes in reverse order to keep indices valid
                    for (int i = indicesToRemove.Count - 1; i >= 0; i--)
                    {
                        sheet.Shapes.RemoveAt(indicesToRemove[i]);
                    }

                    // Determine approximate center of the used range
                    int firstRow = sheet.Cells.MinRow;
                    int firstColumn = sheet.Cells.MinColumn;
                    int lastRow = sheet.Cells.MaxRow;
                    int lastColumn = sheet.Cells.MaxColumn;
                    int centerRow = (firstRow + lastRow) / 2;
                    int centerColumn = (firstColumn + lastColumn) / 2;

                    // Add a new WordArt shape as watermark
                    Shape watermark = sheet.Shapes.AddWordArt(
                        PresetWordArtStyle.WordArtStyle1, // preset style
                        watermarkText,                    // text
                        centerRow,                        // upper left row
                        centerColumn,                     // upper left column
                        0,                                // offset from top of the cell (pixels)
                        0,                                // offset from left of the cell (pixels)
                        400,                              // width (pixels)
                        100                               // height (pixels)
                    );

                    // Format the watermark (semi‑transparent, rotated, no border)
                    watermark.Fill.Transparency = 0.5; // 50% transparent
                    watermark.RotationAngle = -45;     // diagonal
                    watermark.Line.Transparency = 1;   // hide border
                }

                // Save the modified workbook (lifecycle rule: save)
                workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Example usage – replace with actual file paths as needed
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";
            string watermarkText = "CONFIDENTIAL";

            WatermarkHelper.RemoveWordArtAndAddWatermark(inputPath, outputPath, watermarkText);
        }
    }
}
