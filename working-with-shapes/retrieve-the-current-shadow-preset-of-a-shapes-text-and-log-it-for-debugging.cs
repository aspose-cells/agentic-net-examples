using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;
using System.Drawing;

namespace AsposeCellsShadowDemo
{
    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape with some text
            Shape shape = sheet.Shapes.AddRectangle(2, 2, 2, 2, 150, 80);
            shape.Text = "Shadow Text";

            // Access the TextOptions of the shape's text and set a shadow preset
            TextOptions textOpts = shape.Characters(0, shape.Text.Length).TextOptions;
            textOpts.Shadow.PresetType = PresetShadowType.OffsetBottom; // example preset

            // Save the workbook
            string filePath = "ShadowPresetDemo.xlsx";
            workbook.Save(filePath, SaveFormat.Xlsx);

            // ---------- Load the workbook and retrieve the shadow preset ----------
            Workbook loadedWorkbook = new Workbook(filePath);
            Shape loadedShape = loadedWorkbook.Worksheets[0].Shapes[0];

            // Get the TextOptions of the loaded shape's text
            TextOptions loadedTextOpts = loadedShape.Characters(0, loadedShape.Text.Length).TextOptions;

            // Retrieve the current shadow preset type
            PresetShadowType currentPreset = loadedTextOpts.Shadow.PresetType;

            // Log the preset type for debugging
            Console.WriteLine("Current text shadow preset: " + currentPreset);
        }
    }
}