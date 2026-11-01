// Title: C# reflection example to enumerate all PresetShadowType enum values in Aspose.Cells
// Common Searches: how to enumerate PresetShadowType enum in Aspose.Cells with C# reflection | C# code to get list of shape shadow preset types using Aspose.Cells | retrieve all PresetShadowType members programmatically in Aspose.Cells drawing API | using Enum.GetNames to display Aspose.Cells PresetShadowType values | dynamic shape shadow configuration by listing PresetShadowType enumeration in C#
// Tags: reflection enumerate drawing enums Aspose.Cells | shape shadow preset enumeration Aspose.Cells | Aspose.Cells PresetShadowType enumeration | C# Enum.GetNames Aspose.Cells drawing | dynamic shape shadow configuration Aspose.Cells

using System;
using Aspose.Cells.Drawing;

// The sample obtains the Type of the PresetShadowType enum from Aspose.Cells.Drawing, uses Enum.GetNames and Enum.GetValues to collect every member name and its integer value, and writes each name‑value pair to the console for dynamic shape shadow configuration.
class Program
{
    static void Main()
    {
        // Get the enum type for PresetShadowType
        Type enumType = typeof(PresetShadowType);

        // Retrieve all enum names and values
        string[] names = Enum.GetNames(enumType);
        Array values = Enum.GetValues(enumType);

        Console.WriteLine("Available PresetShadowType members:");
        for (int i = 0; i < names.Length; i++)
        {
            // Cast the value to int for display
            int intValue = (int)values.GetValue(i);
            Console.WriteLine($"{names[i]} = {intValue}");
        }
    }
}
