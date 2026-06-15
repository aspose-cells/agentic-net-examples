using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShadowDemo
{
    class Program
    {
        // Predefined shadow style to be applied to every picture
        private static readonly PresetShadowType PredefinedShadow = PresetShadowType.OffsetDiagonalBottomRight;

        // Helper method that adds a picture and automatically applies the predefined shadow
        private static Picture AddPictureWithShadow(Worksheet sheet, int row, int column, string picturePath)
        {
            // Verify that the image file exists before attempting to add it
            if (!File.Exists(picturePath))
                throw new FileNotFoundException($"Image file not found: {picturePath}");

            // Add the picture to the worksheet
            int pictureIndex = sheet.Pictures.Add(row, column, picturePath);
            Picture picture = sheet.Pictures[pictureIndex];

            // Apply the predefined shadow effect
            picture.ShadowEffect.PresetType = PredefinedShadow;

            // Additional optional shadow settings can be configured here if needed
            // picture.ShadowEffect.Transparency = 0.3f;
            // picture.ShadowEffect.Size = 100; // size in percent

            return picture;
        }

        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Path to the image file to be inserted
                string imagePath = "example.jpg";

                // Add a picture with the automatic shadow style if the file exists
                if (File.Exists(imagePath))
                {
                    AddPictureWithShadow(worksheet, 2, 2, imagePath);
                }
                else
                {
                    Console.WriteLine($"Warning: Image file '{imagePath}' not found. Skipping picture insertion.");
                }

                // Save the workbook (lifecycle rule: save)
                string outputPath = "WorkbookWithShadowedPictures.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.WriteLine($"File error: {fnfEx.Message}");
            }
            catch (CellsException cellsEx)
            {
                Console.WriteLine($"Aspose.Cells error: {cellsEx.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}