// Title: Toggle ImageActiveXControl Visibility on an Aspose.Cells Chart (C#)
// Description: Creates a workbook, adds a column chart, places an ImageActiveXControl over the chart, loads a PNG into the control, sets PictureSizeMode to Zoom, and uses the IsVisible property to show or hide the picture based on a Boolean flag before saving the file.
// Keywords: Aspose.Cells ImageActiveXControl | C# toggle picture visibility | Excel chart ActiveX image | IsVisible property Aspose.Cells | conditional picture display Excel | Aspose.Cells chart picture control | ActiveX image control visibility | C# Aspose.Cells example
// Common Searches: how to hide ImageActiveXControl on a chart using Aspose.Cells | C# toggle visibility of picture control in Aspose.Cells workbook | set IsVisible property for ActiveX image in Excel with Aspose.Cells | conditional display of chart logo Aspose.Cells .NET | Aspose.Cells example for showing/hiding picture on chart
// Developer Intent: Show or hide an ImageActiveXControl placed on a chart by evaluating a Boolean flag in C# code using Aspose.Cells.
// Use Cases: Display a company logo on a chart only when a reporting flag is enabled. | Hide a watermark image for a clean presentation view. | Toggle a dynamic picture that reflects user input before exporting the workbook.
// AI Prompts: Write C# code with Aspose.Cells that adds an ImageActiveXControl to a chart and toggles its IsVisible property based on a Boolean variable. | Provide an Aspose.Cells example that loads a PNG into an ImageActiveXControl, sets PictureSizeMode to Zoom, and conditionally hides the control before saving the workbook. | Explain how to programmatically change the visibility of an ActiveX picture control on an Excel chart using Aspose.Cells after the workbook is generated.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds a column chart, places an ImageActiveXControl over the chart, loads a PNG into the control, sets PictureSizeMode to Zoom, and uses the IsVisible property to show or hide the picture based on a Boolean flag before saving the file.
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
            // Flag that determines whether the picture control should be visible
            bool showPictureControl = true; // Change to false to hide the control

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data for the chart
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

            // Add an Image ActiveX control (acts as a picture control) over the chart area
            Shape pictureShape = sheet.Shapes.AddActiveXControl(
                ControlType.Image,    // Image ActiveX control
                6,                    // top row (inside chart area)
                1,                    // left column
                0,                    // top offset (pixels)
                0,                    // left offset (pixels)
                100,                  // width (pixels)
                100);                 // height (pixels)

            // Cast the ActiveXControl to ImageActiveXControl to access picture-specific members
            ImageActiveXControl imageControl = (ImageActiveXControl)pictureShape.ActiveXControl;

            // Load picture data if the file exists
            string imagePath = "sample_image.png";
            if (File.Exists(imagePath))
            {
                try
                {
                    byte[] pictureData = File.ReadAllBytes(imagePath);
                    imageControl.Picture = pictureData;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to load image: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine($"Image file '{imagePath}' not found. Skipping picture assignment.");
            }

            // Toggle visibility based on the flag
            imageControl.IsVisible = showPictureControl;

            // Optionally, set other properties (e.g., picture size mode)
            imageControl.PictureSizeMode = ControlPictureSizeMode.Zoom;

            // Save the workbook
            string outputPath = "TogglePictureControlVisibility.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
