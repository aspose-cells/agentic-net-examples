using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

class EmbeddedImageCombineDemo
{
    static void Main()
    {
        try
        {
            // Prepare a temporary PNG image (1x1 pixel) as a byte array
            byte[] pngBytes = Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2ZcAAAAASUVORK5CYII=");

            // Write the image to a temporary file (optional, used for cleanup demonstration)
            string imagePath = Path.Combine(Path.GetTempPath(), "tempImage.png");
            File.WriteAllBytes(imagePath, pngBytes);

            // Create source workbook and embed image in cell B2
            Workbook sourceWb = new Workbook();
            Worksheet sourceWs = sourceWb.Worksheets[0];
            sourceWs.Cells["B2"].EmbeddedImage = pngBytes;
            sourceWs.Name = "Source";

            // Create destination workbook and embed image in cell D4
            Workbook destWb = new Workbook();
            Worksheet destWs = destWb.Worksheets[0];
            destWs.Cells["D4"].EmbeddedImage = pngBytes;
            destWs.Name = "Destination";

            // Verify embedded images before combine
            Console.WriteLine("Source workbook embedded image count: " + CountEmbeddedImages(sourceWs));
            Console.WriteLine("Destination workbook embedded image count: " + CountEmbeddedImages(destWs));

            // Combine source into destination
            destWb.Combine(sourceWb);

            // Save combined workbook
            string combinedPath = Path.Combine(Path.GetTempPath(), "CombinedWorkbook.xlsx");
            destWb.Save(combinedPath, SaveFormat.Xlsx);

            // Reload combined workbook to ensure persistence
            if (File.Exists(combinedPath))
            {
                Workbook combinedWb = new Workbook(combinedPath);
                Worksheet combinedWs = combinedWb.Worksheets[0];

                // Verify embedded images after combine
                Console.WriteLine("Combined workbook embedded image count: " + CountEmbeddedImages(combinedWs));
            }
            else
            {
                Console.WriteLine("Combined workbook file not found: " + combinedPath);
            }

            // Clean up temporary image file
            if (File.Exists(imagePath))
                File.Delete(imagePath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }

    // Helper method to count cells that contain embedded pictures
    static int CountEmbeddedImages(Worksheet ws)
    {
        int count = 0;
        IEnumerator enumerator = ws.Cells.GetCellsWithPlaceInCellPicture();
        while (enumerator.MoveNext())
        {
            Cell cell = (Cell)enumerator.Current;
            if (cell.EmbeddedImage != null && cell.EmbeddedImage.Length > 0)
                count++;
        }
        return count;
    }
}