// Title: How to batch add uniquely named TextBox shapes to every worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add a TextBox to each worksheet in a loop | create TextBox shapes with unique identifiers across multiple sheets using Aspose.Cells | set TextBox placement to MoveAndSize for all worksheets in Aspose.Cells | apply font and fill formatting to TextBox shapes in an Aspose.Cells workbook | add a hyperlink to a TextBox shape with Aspose.Cells in C#
// Tags: add textbox shape to worksheets Aspose.Cells | assign unique textbox name Aspose.Cells C# | textbox placement moveandsize Aspose.Cells | textbox font formatting fill Aspose.Cells | textbox hyperlink Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace BatchAddTextBoxes
{
    // The example loads an existing workbook or creates a new one, then iterates through every worksheet and adds a TextBox at row 2, column 2 with a 200 × 80 pixel size. Each TextBox receives a unique name composed of the sheet index and a GUID, displays custom text, uses Arial bold formatting, is set to MoveAndSize placement, gets a light‑yellow fill, and includes a hyperlink to Aspose.com. Finally, the workbook is saved to the specified output file.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the input workbook (if it exists). If not, a new workbook will be created.
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            Workbook workbook = null;

            try
            {
                if (File.Exists(inputPath))
                {
                    Console.WriteLine($"Loading workbook from \"{inputPath}\"...");
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    Console.WriteLine($"Input file not found. Creating a new workbook with sample worksheets.");
                    workbook = new Workbook();
                    // Add a few extra worksheets for demonstration
                    workbook.Worksheets.Add("Sheet2");
                    workbook.Worksheets.Add("Sheet3");
                }

                // Iterate through each worksheet and add a TextBox
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    Worksheet ws = workbook.Worksheets[i];
                    Console.WriteLine($"Processing worksheet \"{ws.Name}\" (index {i})...");

                    // Define position and size (row, column, width, height) in pixels
                    // Here we place the textbox at row 2, column 2 with a size of 200x80 pixels
                    int textboxIndex = ws.TextBoxes.Add(2, 2, 200, 80);
                    TextBox tb = ws.TextBoxes[textboxIndex];

                    // Assign a unique name (identifier) to the textbox
                    string uniqueName = $"TB_{i}_{Guid.NewGuid():N}";
                    tb.Name = uniqueName;

                    // Set the displayed text
                    tb.Text = $"TextBox on \"{ws.Name}\"";

                    // Optional formatting
                    tb.Font.Name = "Arial";
                    tb.Font.Size = 12;
                    tb.Font.IsBold = true;
                    tb.Font.Color = Color.DarkBlue;

                    // Set placement so the textbox moves and resizes with cells
                    tb.Placement = PlacementType.MoveAndSize;

                    // Fill background color for visibility
                    tb.Fill.SolidFill.Color = Color.LightYellow;

                    // Add a hyperlink to demonstrate additional features (optional)
                    tb.AddHyperlink("https://www.aspose.com");

                    Console.WriteLine($"  Added TextBox \"{uniqueName}\" at row 2, column 2.");
                }

                // Save the modified workbook
                Console.WriteLine($"Saving workbook to \"{outputPath}\"...");
                workbook.Save(outputPath);
                Console.WriteLine("Operation completed successfully.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("An error occurred:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
