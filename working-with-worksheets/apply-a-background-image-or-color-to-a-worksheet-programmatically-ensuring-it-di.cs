using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBackgroundDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // 1. Apply a background image to the worksheet
            // -------------------------------------------------
            string imagePath = "background.jpg";
            if (File.Exists(imagePath))
            {
                // Set the background image using the file bytes
                worksheet.BackgroundImage = File.ReadAllBytes(imagePath);
            }

            // -------------------------------------------------
            // 2. Apply a solid background color to the entire sheet
            // -------------------------------------------------
            Style bgStyle = workbook.CreateStyle();
            bgStyle.Pattern = BackgroundType.Solid;
            bgStyle.BackgroundColor = Color.LightYellow;

            // Apply the style to a large range to cover the sheet
            int maxRow = 100;
            int maxCol = 50;
            for (int row = 0; row < maxRow; row++)
            {
                for (int col = 0; col < maxCol; col++)
                {
                    worksheet.Cells[row, col].SetStyle(bgStyle);
                }
            }

            // -------------------------------------------------
            // Save the workbook
            // -------------------------------------------------
            workbook.Save("WorksheetWithBackground.xlsx");

            Console.WriteLine("Worksheet background image and color applied successfully.");
        }
    }
}