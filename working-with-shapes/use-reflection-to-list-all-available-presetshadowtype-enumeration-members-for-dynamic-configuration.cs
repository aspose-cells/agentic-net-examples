// Title: List all PresetShadowType enumeration members using reflection in C# with Aspose.Cells
// AI Prompts: Write C# code that uses reflection to retrieve every member name of Aspose.Cells.Drawing.PresetShadowType and prints them to the console. | Create a method that obtains the PresetShadowType enum type at runtime, stores its names in a string array, and returns the array for further processing. | Show how to modify the reflection example to select only PresetShadowType members whose names contain a given substring, such as "Outer".
// Common Searches: how to get all PresetShadowType enum values with reflection in Aspose.Cells C# | C# enumerate Aspose.Cells shape shadow presets at runtime | retrieve PresetShadowType member names dynamically using Aspose.Cells API | using Enum.GetNames to list shape shadow types in Aspose.Cells | dynamic configuration of shape shadows based on PresetShadowType enumeration
// Tags: Aspose.Cells PresetShadowType members | shape shadow preset enumeration | Enum.GetNames Aspose.Cells | dynamic shape shadow values | reflection enum listing C#

using System;
using Aspose.Cells.Drawing;

// The program obtains the PresetShadowType enum type from Aspose.Cells.Drawing, uses Enum.GetNames via reflection to collect all member names into a string array, and writes each name to the console for dynamic shape‑shadow configuration.
class Program
{
    static void Main()
    {
        // Get the enum type for PresetShadowType
        Type enumType = typeof(PresetShadowType);

        // Retrieve all enum member names via reflection
        string[] memberNames = Enum.GetNames(enumType);

        // Display the available members
        Console.WriteLine("Available PresetShadowType members:");
        foreach (string name in memberNames)
        {
            Console.WriteLine(name);
        }
    }
}
