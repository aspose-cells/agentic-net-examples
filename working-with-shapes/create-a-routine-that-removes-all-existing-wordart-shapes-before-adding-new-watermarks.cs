// Title: Remove existing WordArt shapes and insert a rotated semi‑transparent WordArt watermark in each worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells delete all WordArt shapes from a worksheet before adding a watermark | Add a diagonal semi‑transparent WordArt watermark to every sheet in an Excel file with Aspose.Cells | Iterate Aspose.Cells shape collection to remove WordArt programmatically | Replace existing WordArt with a new watermark in an Excel workbook using Aspose.Cells .NET | Aspose.Cells sample for clearing WordArt and creating a rotated watermark in C#
// Tags: delete WordArt shapes Aspose.Cells | add rotated WordArt watermark Aspose.Cells | shape collection iteration Aspose.Cells C# | WordArt transparency Aspose.Cells | Excel worksheet watermark automation C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsWordArtWatermark
{
    // The example loads an Excel workbook, loops through each worksheet, removes any WordArt shapes by checking the IsWordArt property, then adds a semi‑transparent, light‑gray WordArt shape rotated -45° as a watermark, and saves the file. If the input file is missing, a simple sample workbook is created before processing.
    class Program
    {
        static void Main(string[] args)
        {
            // Default file names – can be overridden via command‑line arguments.
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";
            string watermarkText = "CONFIDENTIAL";

            if (args.Length >= 1)
                inputPath = args[0];
            if (args.Length >= 2)
                outputPath = args[1];
            if (args.Length >= 3)
                watermarkText = args[2];

            try
            {
                // Ensure the input file exists; if not, create a simple workbook.
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a sample workbook.");
                    CreateSampleWorkbook(inputPath);
                }

                // Load the workbook.
                Workbook workbook = new Workbook(inputPath);

                // Process each worksheet.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    RemoveAllWordArt(sheet);
                    AddWatermarkWordArt(sheet, watermarkText);
                }

                // Save the modified workbook.
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Processing completed. Output saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }

        /// <summary>
        /// Removes every shape that is identified as WordArt from the given worksheet.
        /// </summary>
        /// <param name="worksheet">The worksheet to clean.</param>
        private static void RemoveAllWordArt(Worksheet worksheet)
        {
            ShapeCollection shapes = worksheet.Shapes;

            // Iterate in reverse order so removal does not affect the index of remaining items.
            for (int i = shapes.Count - 1; i >= 0; i--)
            {
                Shape shape = shapes[i];
                // The IsWordArt property is only meaningful for legacy XLS shapes,
                // but it works for any shape that was added as WordArt via the API.
                if (shape.IsWordArt)
                {
                    shapes.RemoveAt(i);
                    Console.WriteLine($"Removed WordArt shape at index {i} on sheet \"{worksheet.Name}\".");
                }
            }
        }

        /// <summary>
        /// Adds a semi‑transparent WordArt watermark to the worksheet.
        /// </summary>
        /// <param name="worksheet">The worksheet where the watermark will be placed.</param>
        /// <param name="text">The watermark text.</param>
        private static void AddWatermarkWordArt(Worksheet worksheet, string text)
        {
            // Position the watermark roughly in the centre of the sheet.
            // Row/column indices are zero‑based. Offsets are in pixels.
            int topRow = 5;
            int topOffset = 0;
            int leftColumn = 2;
            int leftOffset = 0;
            int height = 200;
            int width = 600;

            // Add the WordArt shape.
            Shape watermark = worksheet.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle1,
                text,
                topRow,
                topOffset,
                leftColumn,
                leftOffset,
                height,
                width);

            // Configure appearance: light gray, 30% transparency, rotated.
            watermark.Fill.FillType = FillType.Solid;
            watermark.Fill.SolidFill.Color = System.Drawing.Color.LightGray;
            watermark.Fill.Transparency = 0.7; // 0 = opaque, 1 = fully transparent

            // Rotate the watermark for a typical diagonal look.
            watermark.RotationAngle = -45;

            // Send the watermark to the back so it does not obscure data.
            watermark.ZOrderPosition = 0;

            Console.WriteLine($"Added watermark \"{text}\" to sheet \"{worksheet.Name}\".");
        }

        /// <summary>
        /// Creates a simple workbook with a few cells populated, used when the input file does not exist.
        /// </summary>
        /// <param name="path">File path where the workbook will be saved.</param>
        private static void CreateSampleWorkbook(string path)
        {
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;

            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Price");
            cells["A2"].PutValue("Apple");
            cells["B2"].PutValue(1.20);
            cells["A3"].PutValue("Banana");
            cells["B3"].PutValue(0.80);
            cells["A4"].PutValue("Total");
            cells["B4"].Formula = "=SUM(B2:B3)";

            wb.CalculateFormula();
            wb.Save(path, SaveFormat.Xlsx);
            Console.WriteLine($"Sample workbook created at \"{path}\".");
        }
    }
}
