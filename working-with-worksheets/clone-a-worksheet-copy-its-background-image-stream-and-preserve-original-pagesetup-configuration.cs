using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsWorksheetCloneDemo
{
    class Program
    {
        static void Main()
        {
            // Create a source workbook and add some data
            Workbook sourceWorkbook = new Workbook();
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            sourceSheet.Name = "OriginalSheet";
            sourceSheet.Cells["A1"].PutValue("Original Worksheet");
            sourceSheet.Cells["B2"].PutValue(123);

            // Set a page‑setup configuration on the source sheet
            sourceSheet.PageSetup.PaperSize = PaperSizeType.PaperA3;
            sourceSheet.PageSetup.Orientation = PageOrientationType.Landscape;
            sourceSheet.PageSetup.PrintArea = "A1:B2";

            // Set a background image on the source sheet (optional – replace with a valid path)
            string imagePath = "background.jpg";
            if (File.Exists(imagePath))
            {
                byte[] imageData = File.ReadAllBytes(imagePath);
                sourceSheet.BackgroundImage = imageData;
            }

            // Add a new worksheet to act as the clone
            Worksheet clonedSheet = sourceWorkbook.Worksheets.Add("ClonedSheet");

            // Copy contents and formats from the source worksheet
            clonedSheet.Copy(sourceSheet);

            // Copy the background image stream
            clonedSheet.BackgroundImage = sourceSheet.BackgroundImage;

            // Preserve the original page‑setup configuration
            clonedSheet.PageSetup.Copy(sourceSheet.PageSetup, new CopyOptions());

            // Save the workbook with the cloned worksheet
            sourceWorkbook.Save("ClonedWorksheetDemo.xlsx");
        }
    }
}