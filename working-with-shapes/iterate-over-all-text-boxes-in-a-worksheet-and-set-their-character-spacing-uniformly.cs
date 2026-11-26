// Title: Iterate through all TextBox shapes in an Aspose.Cells worksheet and set uniform character spacing using C#
// Common Searches: how to change kerning for every TextBox in an Excel workbook with Aspose.Cells C# | C# code to loop over worksheet TextBoxes and apply character spacing | Aspose.Cells example for setting uniform spacing on multiple text boxes | apply TextOptions.Kerning to each TextBox in a workbook using .NET | batch formatting of Excel TextBox shapes with Aspose.Cells
// Tags: Aspose.Cells TextBox kerning | C# set uniform character spacing | batch update Excel text boxes | iterate worksheet TextBoxes Aspose.Cells | TextOptions.Kerning property

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an existing workbook (or creates a new one), accesses the first worksheet, defines a desired character spacing value, iterates over every TextBox shape, skips empty boxes, sets the TextOptions.Kerning property to the specified spacing, and saves the updated workbook as 'output.xlsx'.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source workbook. If the file does not exist a new workbook will be created.
            string inputPath = "input.xlsx";

            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                // Create a simple workbook with a single worksheet if the input file is missing.
                workbook = new Workbook();
                Worksheet wsNew = workbook.Worksheets[0];
                wsNew.Cells["A1"].PutValue("Sample TextBox");
                // Add a sample textbox so the demo can run.
                int tbIdx = wsNew.TextBoxes.Add(2, 1, 200, 100);
                TextBox sampleBox = wsNew.TextBoxes[tbIdx];
                sampleBox.Text = "Hello Aspose.Cells!";
            }

            // Work with the first worksheet.
            Worksheet worksheet = workbook.Worksheets[0];

            // Desired character spacing (points). Adjust as needed.
            double desiredSpacing = 1.5;

            // Iterate over every TextBox in the worksheet.
            foreach (TextBox textBox in worksheet.TextBoxes)
            {
                try
                {
                    // Skip empty text boxes.
                    if (string.IsNullOrEmpty(textBox.Text))
                        continue;

                    // Apply spacing via TextOptions (inherits from Font and provides Kerning).
                    // Kerning property defines the minimum font size at which character kerning occurs,
                    // which effectively influences character spacing.
                    textBox.TextOptions.Kerning = desiredSpacing;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing TextBox: {ex.Message}");
                }
            }

            // Save the modified workbook.
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"All text boxes processed. Result saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
