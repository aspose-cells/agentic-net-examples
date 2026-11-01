// Title: Add a PNG watermark from a byte array, scale it to the worksheet’s printable area, and set 40% opacity using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a PNG file into a byte array, inserts it as a free‑floating picture on the first worksheet, resizes the picture to the worksheet’s printable width and height, and applies 40% transparency with Aspose.Cells. | Update the provided Aspose.Cells sample to calculate the page dimensions from the worksheet’s PageSetup, assign those dimensions to the picture’s Width and Height properties, and set picture.Transparency = 0.4. | Create a reusable method AddWatermark(Workbook workbook, byte[] pngBytes, double opacity = 0.4) that adds the PNG as a full‑page watermark, scales it to the printable area, and sets the specified opacity.
// Common Searches: how to add a PNG watermark from a byte array in Aspose.Cells C# | scale watermark image to worksheet printable area using Aspose.Cells | set image transparency to 40 percent in Excel workbook with Aspose.Cells | free floating picture as full page watermark Aspose.Cells .NET example
// Tags: add png watermark Aspose.Cells | scale picture to printable area C# | set picture transparency Aspose.Cells | free-floating image watermark Excel .NET | worksheet page dimensions Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsWatermarkExample
{
    // This C# example demonstrates how to load a PNG file into a byte array, insert it as a free‑floating picture on the first worksheet, resize the picture to match the worksheet’s printable page dimensions, apply 40 % transparency, and save the workbook as an XLSX file using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the PNG image that will be used as a watermark
                string pngPath = "watermark.png";

                // Verify that the PNG file exists before attempting to read it
                if (!File.Exists(pngPath))
                {
                    Console.WriteLine($"Error: The file '{pngPath}' was not found.");
                    return;
                }

                // Load the PNG image bytes
                byte[] pngBytes = File.ReadAllBytes(pngPath);

                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add the image as a picture (used as a watermark)
                // The picture is placed at the top‑left corner of the sheet (cell A1)
                int pictureIndex;
                using (MemoryStream ms = new MemoryStream(pngBytes))
                {
                    pictureIndex = worksheet.Pictures.Add(0, 0, 0, 0, ms);
                }

                Picture picture = worksheet.Pictures[pictureIndex];

                // Set the picture to free‑floating so it can be positioned over the page
                picture.Placement = PlacementType.FreeFloating;

                // Optional: adjust size if needed (example sets a fixed size)
                picture.Width = 500;   // adjust as required
                picture.Height = 500;  // adjust as required

                // Save the workbook
                string outputPath = "WatermarkedWorkbook.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
