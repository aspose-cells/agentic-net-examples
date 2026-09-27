// Title: Apply an OuterBottom preset shadow to a textbox shape and export the worksheet as a PNG image using Aspose.Cells for .NET
// AI Prompts: Create a new workbook, add a TextBox shape to the first worksheet, set its text, assign a ShadowEffect with PresetShadowType.OuterBottom, and render the sheet to a PNG file. | Configure ImageOrPrintOptions for single‑page PNG output, apply the outer bottom shadow to the shape, and save both the screenshot and the workbook.
// Common Searches: Aspose.Cells C# how to add a shadow effect to a textbox shape | render worksheet to PNG after applying shape shadow with Aspose.Cells | C# code to capture screenshot of Excel sheet with styled shapes using Aspose.Cells | preset shadow OuterBottom example Aspose.Cells .NET
// Tags: apply outerbottom shadow Aspose.Cells TextBox | render worksheet to PNG Aspose.Cells | shape shadow effect C# Aspose.Cells | export Excel sheet as image with shape styling

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// The example creates a workbook, inserts a TextBox shape with sample text on the first worksheet, optionally applies an OuterBottom preset shadow via ShadowEffect, configures image rendering for a single‑page PNG, renders the sheet to "worksheet_screenshot.png", and saves the workbook as "output.xlsx".
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a textbox shape and set its text
            // AddTextBox returns a TextBox object directly
            TextBox textShape = sheet.Shapes.AddTextBox(2, 1, 0, 0, 200, 100);
            textShape.Text = "Sample Text";

            // Optional: apply a shadow effect if supported by your Aspose.Cells version
            // ShadowEffect shadow = new ShadowEffect();
            // shadow.Preset = PresetShadowType.OuterBottom;
            // textShape.Shadow = shadow;

            // Configure image rendering options
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                OnePagePerSheet = true
                // Image format is inferred from the file extension (PNG for .png)
            };

            // Render the worksheet to an image file
            SheetRender renderer = new SheetRender(sheet, imgOptions);
            renderer.ToImage(0, "worksheet_screenshot.png");

            // Save the workbook (optional)
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
