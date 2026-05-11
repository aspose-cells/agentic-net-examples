using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEmbeddedImageMergeDemo
{
    class Program
    {
        static void Main()
        {
            // ---------- Step 1: Prepare a source workbook with an embedded image ----------
            Workbook sourceWorkbook = new Workbook();
            // Create a simple 1x1 PNG image (transparent) from Base64 string
            string pngBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ/6VYVAAAAAElFTkSuQmCC";
            byte[] imageBytes = Convert.FromBase64String(pngBase64);
            // Embed the image into cell B3 (PlaceInCell)
            sourceWorkbook.Worksheets[0].Cells["B3"].EmbeddedImage = imageBytes;
            // Save the source workbook to disk (required for later loading if needed)
            string sourcePath = "SourceWithImage.xlsx";
            sourceWorkbook.Save(sourcePath, SaveFormat.Xlsx);

            // ---------- Step 2: Create a destination workbook ----------
            Workbook destWorkbook = new Workbook();
            destWorkbook.Worksheets[0].Cells["A1"].PutValue("Destination Data");

            // ---------- Step 3: Combine the source workbook into the destination ----------
            // Load the source workbook (demonstrates load rule usage)
            Workbook loadedSource = new Workbook(sourcePath);
            destWorkbook.Combine(loadedSource);

            // ---------- Step 4: Save the combined workbook ----------
            string combinedPath = "CombinedWorkbook.xlsx";
            destWorkbook.Save(combinedPath, SaveFormat.Xlsx);

            // ---------- Step 5: Reload the combined workbook and verify embedded images ----------
            Workbook combinedWorkbook = new Workbook(combinedPath);
            Cells cells = combinedWorkbook.Worksheets[0].Cells;

            int embeddedImageCount = 0;
            IEnumerator enumerator = cells.GetCellsWithPlaceInCellPicture();
            while (enumerator != null && enumerator.MoveNext())
            {
                Cell cell = (Cell)enumerator.Current;
                if (cell.EmbeddedImage != null && cell.EmbeddedImage.Length > 0)
                {
                    embeddedImageCount++;
                    Console.WriteLine($"Embedded image found in cell {cell.Name}");
                }
            }

            Console.WriteLine($"Total cells with embedded images after combine: {embeddedImageCount}");
        }
    }
}