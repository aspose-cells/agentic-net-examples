using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

class Program
{
    // Author: Aspose.Cells example – verifies embedded images after Workbook.Combine
    static void Main()
    {
        // Prepare a sample image (ensure the file exists at the specified path)
        string imagePath = "sample.png";
        if (!File.Exists(imagePath))
        {
            Console.WriteLine($"Image file not found: {imagePath}");
            return;
        }
        byte[] imageBytes = File.ReadAllBytes(imagePath);

        // Create first workbook with an embedded image in cell B3
        Workbook wb1 = new Workbook();
        Worksheet ws1 = wb1.Worksheets[0];
        ws1.Cells["B3"].EmbeddedImage = imageBytes;

        // Create second workbook with an embedded image in cell D5
        Workbook wb2 = new Workbook();
        Worksheet ws2 = wb2.Worksheets[0];
        ws2.Cells["D5"].EmbeddedImage = imageBytes;

        // Combine wb2 into wb1
        wb1.Combine(wb2);

        // Save the combined workbook
        string combinedPath = "CombinedWorkbook.xlsx";
        wb1.Save(combinedPath, SaveFormat.Xlsx);

        // Reload the combined workbook to verify persistence
        Workbook reloaded = new Workbook(combinedPath);
        Worksheet combinedSheet = reloaded.Worksheets[0];
        Cells cells = combinedSheet.Cells;

        // Count cells that contain embedded images
        int embeddedImageCount = 0;
        IEnumerator enumerator = cells.GetCellsWithPlaceInCellPicture();
        while (enumerator.MoveNext())
        {
            Cell cell = (Cell)enumerator.Current;
            if (cell.EmbeddedImage != null && cell.EmbeddedImage.Length > 0)
            {
                embeddedImageCount++;
                Console.WriteLine($"Embedded image found in cell {cell.Name}");
            }
        }

        Console.WriteLine($"Total cells with embedded images after combine and reload: {embeddedImageCount}");
    }
}