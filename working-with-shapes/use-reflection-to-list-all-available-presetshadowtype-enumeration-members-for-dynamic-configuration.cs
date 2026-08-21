// Title: Enumerate PresetShadowType enum values via reflection in Aspose.Cells for .NET
// Description: Demonstrates how to initialize Aspose.Cells, obtain the PresetShadowType enum type, and use reflection to list every member name with its integer value. The sample also saves a workbook to show the typical create‑save workflow.
// Keywords: Aspose.Cells | PresetShadowType | enum reflection | C# | .NET | list enum members | retrieve enum values | shape shadow presets | Aspose.Cells Drawing | enumeration introspection
// Common Searches: list all PresetShadowType values Aspose.Cells | C# reflection enum Aspose.Cells Drawing | how to get numeric value of PresetShadowType | enumerate shadow presets for shapes in Aspose.Cells | retrieve Aspose.Cells PresetShadowType members programmatically
// Developer Intent: The developer needs to programmatically obtain every PresetShadowType enumeration member and its underlying integer value using reflection.
// Use Cases: Populate a dropdown or palette with every shadow preset for shape styling in a UI. | Validate user‑provided numeric codes against defined PresetShadowType values before applying them. | Generate documentation or logs that list supported shadow types for debugging or reporting.
// AI Prompts: Create a method that returns a Dictionary<string,int> of all PresetShadowType names and values using reflection. | Show code that applies a user‑selected PresetShadowType to a shape after enumerating the enum. | Provide robust error handling for converting a string to PresetShadowType in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to initialize Aspose.Cells, obtain the PresetShadowType enum type, and use reflection to list every member name with its integer value. The sample also saves a workbook to show the typical create‑save workflow.
class Program
{
    static void Main()
    {
        // Create a workbook (optional, ensures Aspose.Cells is initialized)
        Workbook workbook = new Workbook();

        // Get the enum type for PresetShadowType
        Type presetShadowEnum = typeof(PresetShadowType);

        // List all enum members with their integer values
        Console.WriteLine("Available PresetShadowType members:");
        foreach (string name in Enum.GetNames(presetShadowEnum))
        {
            int value = (int)Enum.Parse(presetShadowEnum, name);
            Console.WriteLine($"{name} = {value}");
        }

        // Save the workbook to demonstrate the typical lifecycle (create, save)
        workbook.Save("ReflectionPresetShadowTypeDemo.xlsx");
    }
}
