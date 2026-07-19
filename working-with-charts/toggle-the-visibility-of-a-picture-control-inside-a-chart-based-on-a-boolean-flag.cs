// Title: Toggle ImageActiveXControl Visibility on an Excel Chart with Aspose.Cells for .NET
// Description: Shows how to create a workbook, add a column chart, overlay an ImageActiveXControl, and control its visibility by setting the IsVisible property based on a Boolean flag before saving the file.
// Keywords: Aspose.Cells | ImageActiveXControl | chart overlay | toggle visibility | .NET | Excel ActiveX image | IsVisible property | C# example
// Common Searches: Aspose.Cells hide image on chart | C# toggle ActiveX picture visibility in Excel | Set IsVisible for ImageActiveXControl Aspose | Show picture overlay in Excel chart using Aspose.Cells | Programmatically control chart picture visibility .NET
// Developer Intent: Programmatically show or hide an ImageActiveXControl placed over an Excel chart using a Boolean flag.
// Use Cases: Generate a report where a warning icon appears on a chart only when a condition is met. | Create an interactive workbook that can reveal or conceal picture overlays on multiple charts at runtime. | Automate dashboard generation that toggles visual cues on charts based on data thresholds.
// AI Prompts: Write C# code with Aspose.Cells to add an ImageActiveXControl over a chart and set its IsVisible property from a method parameter. | Explain how to replace the placeholder byte array with an actual image file when displaying the control. | Provide a sample that toggles visibility of several ImageActiveXControls on different charts within the same workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, add a column chart, overlay an ImageActiveXControl, and control its visibility by setting the IsVisible property based on a Boolean flag before saving the file.
    public class TogglePictureInChart
    {
        // Call this method with true to show the picture, false to hide it
        public static void Run(bool showPicture)
        {
            try
            {
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

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Add an Image ActiveX control (acts as a picture) positioned over the chart area
                // Parameters: control type, upper left row, upper left column, top offset, left offset, width, height
                Shape shape = sheet.Shapes.AddActiveXControl(
                    ControlType.Image, // Image control can hold a picture
                    6,   // row
                    1,   // column
                    0,   // top offset
                    0,   // left offset
                    200, // width
                    150  // height
                );

                // Cast the control to ImageActiveXControl to access its properties
                ImageActiveXControl imageControl = (ImageActiveXControl)shape.ActiveXControl;

                // Set picture data (empty byte array used here; replace with actual image bytes as needed)
                imageControl.Picture = new byte[0];

                // Toggle visibility based on the provided flag
                imageControl.IsVisible = showPicture;

                // Save the workbook
                workbook.Save("TogglePictureChart.xlsx", SaveFormat.Xlsx);
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the console application
        public static void Main(string[] args)
        {
            // Example usage: show the picture
            Run(true);
        }
    }
}
