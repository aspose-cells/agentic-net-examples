using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        Shape shape = sheet.Shapes.AddRectangle(1, 0, 1, 0, 150, 100);

        // Change the shadow preset to PerspectiveDiagonalLowerLeft
        shape.ShadowEffect.PresetType = PresetShadowType.PerspectiveDiagonalLowerLeft;

        // Save the workbook with the applied shadow effect
        string filePath = "ShadowPresetDemo.xlsx";
        workbook.Save(filePath);

        // Load the saved workbook to verify the shadow preset
        Workbook loadedWorkbook = new Workbook(filePath);
        Shape loadedShape = loadedWorkbook.Worksheets[0].Shapes[0];
        Console.WriteLine("Loaded shadow preset: " + loadedShape.ShadowEffect.PresetType);
    }
}