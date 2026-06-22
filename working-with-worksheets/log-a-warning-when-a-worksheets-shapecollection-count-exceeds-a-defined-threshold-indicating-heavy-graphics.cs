using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Simple program to demonstrate warning when a worksheet has many shapes
    public class HeavyGraphicsWarningDemo
    {
        // Define the threshold for heavy graphics (number of shapes)
        private const int ShapeCountThreshold = 5;

        public static void Main()
        {
            try
            {
                // ---------- Create a new workbook (lifecycle create rule) ----------
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "GraphicsSheet";

                // Add several shapes to the worksheet to exceed the threshold
                sheet.Shapes.AddRectangle(1, 0, 0, 100, 100, 100);
                sheet.Shapes.AddOval(2, 0, 150, 100, 100, 100);
                sheet.Shapes.AddLine(3, 0, 300, 100, 100, 100);
                sheet.Shapes.AddButton(4, 0, 450, 100, 100, 100);
                sheet.Shapes.AddTextBox(5, 0, 600, 100, 100, 100);

                // Add a picture if the file exists
                const string picturePath = "example.jpg";
                if (File.Exists(picturePath))
                {
                    using (FileStream pictureStream = File.OpenRead(picturePath))
                    {
                        // Insert picture at row 0, column 10 (adjust as needed)
                        sheet.Pictures.Add(0, 10, pictureStream);
                    }
                }
                else
                {
                    Console.WriteLine($"Warning: Picture file \"{picturePath}\" not found. Skipping picture insertion.");
                }

                // ---------- Check shape count and log warning if it exceeds the threshold ----------
                int shapeCount = sheet.Shapes.Count;
                if (shapeCount > ShapeCountThreshold)
                {
                    Console.WriteLine($"Warning: Worksheet \"{sheet.Name}\" contains {shapeCount} shapes, which exceeds the threshold of {ShapeCountThreshold}. Consider reducing graphics for better performance.");
                }

                // ---------- Save the workbook (lifecycle save rule) ----------
                workbook.Save("HeavyGraphicsWarningDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}