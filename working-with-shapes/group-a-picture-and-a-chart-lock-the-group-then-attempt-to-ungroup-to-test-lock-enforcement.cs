using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsGroupLockDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // Add a picture to the worksheet (if the file exists)
                // -------------------------------------------------
                Picture? picture = null;
                const string imagePath = "example.jpg";

                if (File.Exists(imagePath))
                {
                    int pictureIndex = sheet.Pictures.Add(0, 0, imagePath);
                    picture = sheet.Pictures[pictureIndex];
                }
                else
                {
                    Console.WriteLine($"Image file \"{imagePath}\" not found. Skipping picture addition.");
                }

                // -------------------------------------------------
                // Add a chart to the worksheet
                // -------------------------------------------------
                // Add returns the chart index; retrieve the Chart object afterwards
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 0);
                Chart chart = sheet.Charts[chartIndex];
                // Set chart size
                chart.ChartObject.Width = 400;
                chart.ChartObject.Height = 300;

                // -------------------------------------------------
                // Group the picture (if added) and the chart shape
                // -------------------------------------------------
                Shape[] shapesToGroup = picture != null
                    ? new Shape[] { picture, chart.ChartObject }
                    : new Shape[] { chart.ChartObject };

                GroupShape group = sheet.Shapes.Group(shapesToGroup);

                // -------------------------------------------------
                // Lock the group (prevent ungrouping when sheet is protected)
                // -------------------------------------------------
                group.SetLockedProperty(ShapeLockType.Group, true);
                group.IsLocked = true;

                // -------------------------------------------------
                // Protect the worksheet to enforce the lock
                // -------------------------------------------------
                sheet.Protection.AllowEditingObject = false;
                sheet.Protect(ProtectionType.All);

                // -------------------------------------------------
                // Attempt to ungroup the locked group and capture the result
                // -------------------------------------------------
                try
                {
                    group.Ungroup();
                    Console.WriteLine("Ungroup succeeded (lock not enforced).");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Ungroup failed as expected: {ex.Message}");
                }

                // -------------------------------------------------
                // Save the workbook
                // -------------------------------------------------
                const string outputPath = "GroupLockTest.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to \"{outputPath}\".");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}