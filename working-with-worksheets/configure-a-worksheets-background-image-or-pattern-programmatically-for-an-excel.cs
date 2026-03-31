using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

class WorksheetBackgroundDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // ----- Set worksheet background image -----
        // Provide the path to an existing image file
        string imagePath = "background.jpg";
        if (File.Exists(imagePath))
        {
            // Read image bytes and assign to the worksheet
            byte[] imageData = File.ReadAllBytes(imagePath);
            worksheet.BackgroundImage = imageData;
        }

        // ----- Set a cell background pattern -----
        // Create a style with a diagonal stripe pattern
        Style style = workbook.CreateStyle();
        style.Pattern = BackgroundType.DiagonalStripe;
        style.ForegroundColor = Color.LightBlue;
        style.BackgroundColor = Color.DarkBlue;

        // Apply the style to cell B2
        Cell cell = worksheet.Cells["B2"];
        cell.PutValue("Pattern Demo");
        cell.SetStyle(style);

        // Save the workbook
        workbook.Save("WorksheetBackgroundDemo.xlsx");
    }
}