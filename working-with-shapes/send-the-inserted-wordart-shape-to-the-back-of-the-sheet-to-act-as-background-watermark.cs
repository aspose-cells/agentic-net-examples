// Title: Insert a WordArt text effect as a background watermark and send it to the back of an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a WordArt shape with custom text, sets its transparency, and moves it behind all other objects in an Excel sheet using Aspose.Cells. | Show how to add a semi‑transparent WordArt watermark to a workbook, adjust its Z‑order, and save the file with Aspose.Cells for .NET. | Provide a sample that loads an existing .xlsx (or creates a new workbook), inserts a text‑effect shape as a background, and positions it at ZOrderPosition 0 using Aspose.Cells.
// Common Searches: how to add a WordArt watermark to an Excel file with Aspose.Cells C# | Aspose.Cells set shape ZOrderPosition to back of sheet | make WordArt shape transparent in Aspose.Cells .NET | create background text effect shape in Excel using Aspose.Cells API | load workbook or create new and add watermark with Aspose.Cells C#
// Tags: add WordArt shape Aspose.Cells .NET | set shape transparency Aspose.Cells | adjust ZOrderPosition Aspose.Cells | background watermark Excel C# | text effect shape Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsWatermark
{
    // // Loads or creates an Excel workbook, adds a WordArt text‑effect shape with the text "CONFIDENTIAL", sets its fill transparency to 70%, moves the shape to the back by setting ZOrderPosition = 0, and saves the result as output.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook or load an existing one if the file is present
                Workbook workbook;
                string inputPath = "input.xlsx";
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                }

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add WordArt (text effect) to serve as a watermark
                // Height and width parameters are required by the API
                Shape wordArt = worksheet.Shapes.AddTextEffect(
                    MsoPresetTextEffect.TextEffect1, // preset effect
                    "CONFIDENTIAL",                  // watermark text
                    "Arial",                         // font name
                    72,                              // font size
                    false,                           // bold
                    false,                           // italic
                    5,                               // upper-left row
                    1,                               // upper-left column
                    30,                              // lower-right row
                    10,                              // lower-right column
                    200,                             // height (points)
                    800                              // width (points)
                );

                // Make the WordArt semi‑transparent using the Fill property
                wordArt.Fill.Transparency = 0.7; // 0 = opaque, 1 = fully transparent

                // Send the shape to the back by setting its Z‑order position
                wordArt.ZOrderPosition = 0;

                // Save the workbook with the watermark applied
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
