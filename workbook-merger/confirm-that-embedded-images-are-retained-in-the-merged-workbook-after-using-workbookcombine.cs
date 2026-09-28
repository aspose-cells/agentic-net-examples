// Title: Check that cell-embedded PNG images survive Workbook.Combine merging in Aspose.Cells for .NET
// AI Prompts: Generate C# code that embeds a PNG into a worksheet cell, merges another workbook with Workbook.Combine, and validates the image remains. | Show how to load a source workbook containing an embedded image, combine it into a destination workbook, and programmatically verify the EmbeddedImage property. | Provide a step-by-step C# example that creates a temporary workbook with an embedded picture, uses Workbook.Combine, and asserts image retention after saving.
// Common Searches: Aspose.Cells keep embedded picture in cell after using Workbook.Combine in C# | C# verify embedded image is retained when merging Excel files with Aspose.Cells | How to preserve cell-placed images during workbook combine operation Aspose.Cells .NET | Workbook.Combine does not copy cell EmbeddedImage property Aspose.Cells example | Check EmbeddedImage after combining workbooks using Aspose.Cells API
// Tags: Workbook.Combine embedded image support | cell EmbeddedImage handling Aspose.Cells | C# merge workbooks while keeping pictures | Aspose.Cells embed PNG in Excel cell | check image presence after combine

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEmbeddedImageCombineDemo
{
    // The sample creates a source workbook with a PNG embedded in cell B3, saves it, then merges this workbook into a destination workbook using Workbook.Combine. After saving the combined file, the code reloads it and reads the EmbeddedImage property of cell B3 to confirm the picture was retained, demonstrating that cell-embedded images survive the combine operation.
    class Program
    {
        static void Main()
        {
            // Path for temporary files
            string sourcePath = "SourceWorkbook.xlsx";
            string combinedPath = "CombinedWorkbook.xlsx";

            // Small 1x1 pixel PNG (base64 encoded)
            string pngBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XcZcAAAAASUVORK5CYII=";
            byte[] imageData = Convert.FromBase64String(pngBase64);

            // ---------- Create source workbook with embedded image ----------
            Workbook sourceWorkbook = new Workbook(); // create rule
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

            // Embed image into cell B3 (place-in-cell picture)
            sourceSheet.Cells["B3"].EmbeddedImage = imageData;

            // Save source workbook (save rule)
            sourceWorkbook.Save(sourcePath, SaveFormat.Xlsx);
            sourceWorkbook.Dispose();

            // ---------- Create destination workbook ----------
            Workbook destWorkbook = new Workbook(); // create rule
            Worksheet destSheet = destWorkbook.Worksheets[0];
            destSheet.Cells["A1"].PutValue("Destination Data");

            // Combine source workbook into destination workbook
            using (Workbook toCombine = new Workbook(sourcePath)) // load rule
            {
                destWorkbook.Combine(toCombine);
            }

            // Save the combined workbook
            destWorkbook.Save(combinedPath, SaveFormat.Xlsx);
            destWorkbook.Dispose();

            // ---------- Verify that the embedded image is retained ----------
            Workbook verifyWorkbook = new Workbook(combinedPath); // load rule
            Worksheet verifySheet = verifyWorkbook.Worksheets[0];
            byte[] embedded = verifySheet.Cells["B3"].EmbeddedImage;

            if (embedded != null && embedded.Length > 0)
            {
                Console.WriteLine("Embedded image is retained after Combine.");
            }
            else
            {
                Console.WriteLine("Embedded image was NOT retained after Combine.");
            }

            // Clean up temporary files (optional)
            // File.Delete(sourcePath);
            // File.Delete(combinedPath);

            verifyWorkbook.Dispose();
        }
    }
}
