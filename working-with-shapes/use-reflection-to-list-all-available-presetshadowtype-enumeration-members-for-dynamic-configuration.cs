using System;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Get the enum type via reflection
        Type presetShadowType = typeof(PresetShadowType);

        // List all enum names and their integer values
        Console.WriteLine("Available PresetShadowType members:");
        foreach (string name in Enum.GetNames(presetShadowType))
        {
            int value = (int)Enum.Parse(presetShadowType, name);
            Console.WriteLine($"{name} = {value}");
        }

        // Demonstrate using a reflected enum value on a shape
        Workbook workbook = new Workbook();                     // create workbook
        Worksheet worksheet = workbook.Worksheets[0];           // access first worksheet
        Shape shape = worksheet.Shapes.AddRectangle(0, 0, 0, 0, 100, 100); // add a rectangle shape

        // Set a shadow preset using the enum name obtained via reflection
        shape.ShadowEffect.PresetType = (PresetShadowType)Enum.Parse(presetShadowType, "OffsetBottom");
        Console.WriteLine($"Shape shadow preset set to: {shape.ShadowEffect.PresetType}");

        // Save the workbook (optional, demonstrates lifecycle usage)
        workbook.Save("ReflectionPresetShadowDemo.xlsx");
    }
}