// Title: Rotate a rectangle shape 30° around the Z‑axis to simulate a sparkline group at cell G1 using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that adds a rectangle shape over cell G1, sets its ThreeDFormat.RotationZ to 30 degrees, and saves the workbook. | Demonstrate how to apply a 30‑degree Z‑axis 3‑D rotation to a shape that represents a sparkline in an Excel file using Aspose.Cells .NET.
// Common Searches: Aspose.Cells C# rotate shape 30 degrees Z axis | how to set ThreeDFormat.RotationZ for a shape in Aspose.Cells | apply 3D rotation to a sparkline placeholder Excel Aspose.Cells .NET | C# Aspose.Cells add rectangle shape over cell G1 and rotate | example of ThreeDFormat rotationz property in Aspose.Cells
// Tags: Aspose.Cells shape threeD rotationz | C# Aspose.Cells add rectangle shape | Excel shape 30 degree Z axis rotation | Aspose.Cells 3D formatting example | simulate sparkline with shape Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace Sparkline3DRotationDemo
{
    // The sample creates a new workbook, inserts a rectangle shape over cell G1 to represent a sparkline, applies a 30‑degree Z‑axis rotation via the ThreeDFormat.RotationZ property, and saves the file as Sparkline3DRotationDemo.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data (used only for demonstration; no sparkline API)
                sheet.Cells["A1"].PutValue(5);
                sheet.Cells["B1"].PutValue(3);
                sheet.Cells["C1"].PutValue(8);
                sheet.Cells["D1"].PutValue(2);

                // Demonstrate 3‑D effect by adding a rectangle shape over the area where a sparkline would be
                // Parameters: type, upper left row, upper left column, lower right row, lower right column, width, height
                Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 0, 6, 0, 6, 20, 100);
                ThreeDFormat threeD = shape.ThreeDFormat;
                threeD.RotationZ = 30; // Rotate the shape 30 degrees around the Z‑axis

                // Prepare output path
                string outputFile = "Sparkline3DRotationDemo.xlsx";
                string outputFullPath = Path.GetFullPath(outputFile);
                string outputDir = Path.GetDirectoryName(outputFullPath);

                // Create directory if needed (outputDir can be null when only a file name is provided)
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputFullPath);
                Console.WriteLine($"Workbook saved successfully to '{outputFullPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
