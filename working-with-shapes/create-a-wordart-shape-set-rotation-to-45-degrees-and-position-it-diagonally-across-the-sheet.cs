// Title: Add a rotated WordArt (TextEffect) shape and position it diagonally on an Excel worksheet using Aspose.Cells for .NET
// Common Searches: how to insert a rotated WordArt shape in an Excel file with Aspose.Cells C# | Aspose.Cells C# create TextEffect shape and set 45 degree rotation | place a WordArt shape diagonally across a worksheet using Aspose.Cells | set rotation angle of a shape in Aspose.Cells workbook | adjust position and size of TextEffect shape in Aspose.Cells
// Tags: add TextEffect shape Aspose.Cells | rotate WordArt shape C# | position shape diagonally Excel | set shape size Aspose.Cells | apply rotation angle Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // The example creates a new workbook, adds a TextEffect (WordArt) shape with custom text and formatting, rotates it 45 degrees, positions it at the top‑left corner to appear diagonal across the sheet, sets its dimensions, and saves the file as WordArt.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Define the bounding rectangle for the WordArt shape
                int upperLeftRow = 0;
                int upperLeftColumn = 0;
                int lowerRightRow = 20;
                int lowerRightColumn = 10;

                // Add a WordArt (TextEffect) shape to the worksheet
                Shape wordArt = worksheet.Shapes.AddTextEffect(
                    MsoPresetTextEffect.TextEffect1,   // preset WordArt style
                    "Aspose.Cells WordArt",            // displayed text
                    "Arial",                           // font name
                    48,                                // font size
                    true,                              // bold
                    false,                             // italic
                    upperLeftRow,
                    upperLeftColumn,
                    lowerRightRow,
                    lowerRightColumn,
                    800,                               // width in points (initial)
                    400);                              // height in points (initial)

                // Rotate the shape 45 degrees
                wordArt.RotationAngle = 45;

                // Position the shape at the sheet origin
                wordArt.Top = 0;
                wordArt.Left = 0;

                // Adjust size if needed
                wordArt.Width = 800;
                wordArt.Height = 400;

                // Save the workbook
                string outputPath = "WordArt.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
