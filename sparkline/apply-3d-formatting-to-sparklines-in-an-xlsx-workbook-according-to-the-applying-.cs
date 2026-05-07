using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsSparklines3D
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the sparkline (5 rows, 4 columns)
            for (int row = 0; row < 5; row++)
            {
                for (int col = 0; col < 4; col++)
                {
                    sheet.Cells[row, col].PutValue(row * 4 + col + 1);
                }
            }

            // Define the location where the sparklines will be placed (column E)
            CellArea location = new CellArea
            {
                StartRow = 0,
                EndRow = 4,
                StartColumn = 4,
                EndColumn = 4
            };

            // Add a sparkline group of type Line that uses the data range A1:D5
            int groupIndex = sheet.SparklineGroups.Add(
                SparklineType.Line,
                "A1:D5",
                false,
                location);

            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // Add sparklines for each row (the Add method creates them automatically,
            // but we call it explicitly to illustrate the process)
            for (int i = 0; i < 5; i++)
            {
                // Data range for each sparkline is the whole row A:D
                string dataRange = $"A{i + 1}:D{i + 1}";
                group.Sparklines.Add(dataRange, i, 4);
            }

            // -----------------------------------------------------------------
            // Apply 3‑D formatting.
            // Sparklines themselves do not expose a ThreeDFormat property,
            // but each sparkline is rendered as a shape internally.
            // We can retrieve the shape that represents the sparkline
            // via the worksheet's Shapes collection. The shape name follows
            // the pattern "Sparkline_{row}_{column}" (e.g., "Sparkline_0_4").
            // -----------------------------------------------------------------
            for (int i = 0; i < 5; i++)
            {
                // Construct the expected shape name for the sparkline at row i, column 4
                string shapeName = $"Sparkline_{i}_{4}";
                Shape sparkShape = sheet.Shapes[shapeName];

                if (sparkShape != null)
                {
                    // Access the ThreeDFormat object of the shape
                    ThreeDFormat threeD = sparkShape.ThreeDFormat;

                    // Example 3‑D settings
                    threeD.ExtrusionColor.Color = Color.LightGray;   // extrusion color
                    threeD.ExtrusionHeight = 5;                     // height in points
                    threeD.LightAngle = 45;                         // light angle
                    threeD.RotationX = 15;                          // rotate around X axis
                    threeD.RotationY = 30;                          // rotate around Y axis
                    threeD.PresetCameraType = PresetCameraType.PerspectiveAbove; // camera view
                }
            }

            // Save the workbook as XLSX
            workbook.Save("SparklinesWith3D.xlsx", SaveFormat.Xlsx);
        }
    }
}