// Title: List all PresetShadowType enum values for Aspose.Cells shapes using C# reflection
// Common Searches: how to retrieve PresetShadowType enum members with reflection in Aspose.Cells C# | enumerate shape shadow preset types at runtime using Aspose.Cells API | C# code to display PresetShadowType names and numeric values for Excel shapes | using Enum.GetNames to get Aspose.Cells shadow type list for dynamic styling | list available shadow presets for shapes in Aspose.Cells via reflection
// Tags: C# reflection enumerate PresetShadowType | Aspose.Cells shape shadow enumeration | Enum.GetNames Aspose.Cells shadow types | dynamic shape shadow configuration C# | list PresetShadowType values Aspose.Cells

using System;
using System.Reflection;
using Aspose.Cells.Drawing;   // Namespace containing PresetShadowType

// The example uses C# reflection to obtain the Type of the PresetShadowType enumeration, verifies it is an enum, retrieves all member names with Enum.GetNames, parses each to its integer value, and prints the name‑value pairs, enabling dynamic configuration of shape shadow presets in Aspose.Cells.
class Program
{
    static void Main()
    {
        // Get the Type object for the PresetShadowType enumeration
        Type shadowEnumType = typeof(PresetShadowType);

        // Ensure the type is an enum
        if (!shadowEnumType.IsEnum)
        {
            Console.WriteLine("PresetShadowType is not an enumeration.");
            return;
        }

        // Retrieve all enum names using reflection
        string[] enumNames = Enum.GetNames(shadowEnumType);

        // Output the available PresetShadowType members
        Console.WriteLine("Available PresetShadowType members:");
        foreach (string name in enumNames)
        {
            // Optionally, get the underlying numeric value
            object value = Enum.Parse(shadowEnumType, name);
            Console.WriteLine($"{name} = {(int)value}");
        }
    }
}
