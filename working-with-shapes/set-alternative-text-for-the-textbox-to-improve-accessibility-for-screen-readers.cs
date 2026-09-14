// Title: Add a textbox shape with AlternativeText for screen‑reader accessibility using Aspose.Cells in C#
// AI Prompts: Create an Excel workbook in C# and insert a textbox shape whose AlternativeText property describes the content for screen readers using Aspose.Cells. | Generate a .xlsx file with a textbox shape that includes accessible alternative text via the Aspose.Cells .NET API. | Set the AlternativeText of a TextBox shape in a worksheet to improve accessibility for assistive technologies with Aspose.Cells.
// Common Searches: how to set alternative text for a textbox shape in Aspose.Cells C# | Aspose.Cells .NET add accessible textbox to Excel worksheet | C# Aspose.Cells AlternativeText property example for shapes | make Excel textbox readable by screen readers using Aspose.Cells | sample code for textbox accessibility in Aspose.Cells workbook
// Tags: Aspose.Cells set textbox AlternativeText | C# add textbox shape accessibility | Aspose.Cells shape alternative text .NET | Excel textbox screen reader support Aspose | Aspose.Cells workbook save accessible shape

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example creates a new Workbook, adds a TextBox shape to the first worksheet, assigns descriptive alternative text to the shape for screen‑reader accessibility, ensures the output directory exists, saves the workbook as AccessibleTextbox.xlsx, and handles any exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook(); // lifecycle: create
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a textbox shape to the worksheet.
                // Overload: AddTextBox(int upperLeftRow, int upperLeftColumn, int lowerRightRow, int lowerRightColumn, int height, int width)
                TextBox textBox = worksheet.Shapes.AddTextBox(
                    5,   // upperLeftRow
                    5,   // upperLeftColumn
                    10,  // lowerRightRow
                    10,  // lowerRightColumn
                    200, // height (in points)
                    100  // width (in points)
                );

                // Set alternative text for accessibility (screen readers)
                textBox.AlternativeText = "Summary of the chart data displayed in this textbox.";

                // Save the workbook
                string outputPath = "AccessibleTextbox.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath); // lifecycle: save

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
