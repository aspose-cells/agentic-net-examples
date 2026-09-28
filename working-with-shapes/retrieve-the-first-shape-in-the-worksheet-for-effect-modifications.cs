// Title: Retrieve the first shape in an Excel worksheet and apply an outer shadow effect with Aspose.Cells for .NET (C#)
// AI Prompts: Load an Excel workbook with Aspose.Cells, obtain the first Shape object from the first worksheet, and enable its ShadowEffect with custom color, transparency, size, blur, distance, and direction. | Use dynamic typing in C# to access a Shape's ShadowEffect property and configure an outer shadow programmatically with Aspose.Cells. | Save the modified workbook to a new file after updating the first shape's shadow settings using Aspose.Cells.
// Common Searches: C# Aspose.Cells get first shape from worksheet and set shadow properties | how to add outer shadow to Excel shape using Aspose.Cells .NET | using dynamic to modify shape ShadowEffect in Aspose.Cells example | apply custom shadow color transparency blur to shape in Excel with Aspose.Cells | retrieve shape collection index 0 and change shadow in Aspose.Cells C#
// Tags: retrieve first shape Aspose.Cells C# | set shape shadow effect Aspose.Cells | dynamic ShadowEffect handling Aspose.Cells | modify Excel shape appearance .NET | apply outer shadow to worksheet shape

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing Excel workbook, checks for shapes on the first worksheet, retrieves the first shape, accesses its ShadowEffect via a dynamic reference, configures visibility, color, transparency, size, blur, distance, and direction to create an outer shadow, and then saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one shape in the worksheet
            if (sheet.Shapes.Count > 0)
            {
                try
                {
                    // Retrieve the first shape
                    Shape shape = sheet.Shapes[0];

                    // Obtain the shape's shadow effect object (using dynamic to handle version differences)
                    dynamic shadow = shape.ShadowEffect;

                    // Apply an outer shadow effect
                    shadow.Visible = true;                 // visibility
                    shadow.Color = Color.Black;            // shadow color
                    shadow.Transparency = 0.5;             // 50% transparent
                    shadow.Size = 5;                       // size of the shadow
                    shadow.Blur = 5;                       // blur radius
                    shadow.Distance = 5;                   // distance from the shape
                    shadow.Direction = 45;                // direction in degrees
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Unable to apply shadow effect. {ex.Message}");
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
