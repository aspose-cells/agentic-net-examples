using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

namespace AsposeCellsExamples
{
    public class TogglePictureControlVisibility
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Flag that determines visibility of the picture control
            bool showPictureControl = true; // set to false to hide

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Create a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Add an Image ActiveX control (acts as a picture control)
            // Position it roughly inside the chart area (rows 6‑8, columns 2‑4)
            var pictureShape = sheet.Shapes.AddActiveXControl(
                ControlType.Image,   // Image control displays a picture
                6,   // top row
                2,   // left column
                0,   // top offset (pixels)
                0,   // left offset (pixels)
                100, // width (pixels)
                60   // height (pixels)
            );

            // Cast to ImageActiveXControl to access picture‑related members
            ImageActiveXControl imgControl = (ImageActiveXControl)pictureShape.ActiveXControl;

            // Load picture data if the file exists
            string imagePath = "sample_image.png";
            if (File.Exists(imagePath))
            {
                byte[] imgData = File.ReadAllBytes(imagePath);
                imgControl.Picture = imgData;
                imgControl.PictureSizeMode = ControlPictureSizeMode.Zoom;
            }
            else
            {
                Console.WriteLine($"Warning: Image file '{imagePath}' not found. Skipping picture assignment.");
            }

            // Toggle visibility based on the flag
            imgControl.IsVisible = showPictureControl;

            // Save the workbook
            string outputPath = "ChartWithToggledPictureControl.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}