// Title: C# – Apply Plastic Material to a Shape Containing a Sparkline with Aspose.Cells
// Description: Demonstrates how to create a workbook, add a line‑type sparkline, insert a rectangle shape next to it, and configure the shape's ThreeDFormat to use the Plastic material, slight extrusion, rotation, perspective camera, and bright‑room lighting for a subtle glossy 3‑D effect, then save the file as XLSX.
// Keywords: Aspose.Cells C# sparkline shape material | Plastic material ThreeDFormat | glossy 3D shape Aspose.Cells | Excel rectangle shape extrusion | sparkline visual styling .NET | PresetMaterialType Plastic | LightRigType BrightRoom | PerspectiveFront camera Aspose | ThreeDFormat lighting and rotation | Aspose.Cells shape formatting
// Common Searches: set shape material to plastic Aspose.Cells C# | add sparkline and 3D rectangle shape in Excel using Aspose | how to apply glossy effect to a shape around a sparkline | Aspose.Cells ThreeDFormat plastic material example | C# code for sparkline with extruded shape
// Developer Intent: Apply the Plastic material to a rectangle shape that frames a sparkline, creating a subtle glossy 3‑D appearance.
// Use Cases: Enhance financial dashboards by framing sparklines with a glossy plastic‑styled shape for better visual emphasis. | Generate automated Excel reports where each trend line is highlighted by a 3‑D rectangle with plastic material. | Design Excel templates that use plastic‑material shapes to give sparklines a modern, polished look.
// AI Prompts: Write C# code using Aspose.Cells to add a line sparkline, insert a rectangle shape, and set its ThreeDFormat.Material to Plastic with extrusion, rotation, and bright‑room lighting. | Show an example that configures a shape's ThreeDFormat for a glossy Plastic material around a sparkline and saves the workbook as XLSX. | Explain step‑by‑step how to change a shape's material to Plastic in Aspose.Cells and how it impacts the rendering of a sparkline.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add a line‑type sparkline, insert a rectangle shape next to it, and configure the shape's ThreeDFormat to use the Plastic material, slight extrusion, rotation, perspective camera, and bright‑room lighting for a subtle glossy 3‑D effect, then save the file as XLSX.
    public class SparklineMaterialPlasticDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // 1. Add sample data that will be used for the sparkline
                sheet.Cells["A1"].PutValue(5);
                sheet.Cells["B1"].PutValue(2);
                sheet.Cells["C1"].PutValue(1);
                sheet.Cells["D1"].PutValue(3);

                // 2. Create a sparkline group (Line type) for the data range A1:D1
                CellArea sparklineLocation = new CellArea
                {
                    StartRow = 0,
                    EndRow = 0,
                    StartColumn = 4,
                    EndColumn = 4
                };
                int sparklineGroupIdx = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, sparklineLocation);
                SparklineGroup sparklineGroup = sheet.SparklineGroups[sparklineGroupIdx];
                sparklineGroup.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);

                // 3. Add a rectangle shape positioned near the sparkline
                // Parameters: type, upperLeftRow, upperLeftColumn, top, left, height, width
                Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 5, 5, 0, 0, 30, 200);
                shape.Text = "Sparkline with Plastic Material";

                // Access the ThreeDFormat of the shape and set material to Plastic
                ThreeDFormat threeD = shape.ThreeDFormat;
                threeD.Material = PresetMaterialType.Plastic;
                threeD.ExtrusionHeight = 5;               // slight depth
                threeD.RotationX = 10;                    // tilt
                threeD.PresetCameraType = PresetCameraType.PerspectiveFront;
                threeD.Lighting = LightRigType.BrightRoom; // good lighting for glossy look

                // 4. Save the workbook
                workbook.Save("SparklineMaterialPlasticDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            SparklineMaterialPlasticDemo.Run();
        }
    }
}
