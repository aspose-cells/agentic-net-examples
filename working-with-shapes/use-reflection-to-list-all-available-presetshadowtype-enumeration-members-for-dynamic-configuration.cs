// Title: C# Example: Use Reflection to Enumerate All PresetShadowType Values in Aspose.Cells .NET
// Description: Shows how to obtain the PresetShadowType enumeration from Aspose.Cells.Drawing via reflection, retrieve each member name and its integer value with Enum.GetNames/GetValues, and print the list to the console. Useful for dynamic UI controls, validation, or documentation of shape shadow presets.
// Keywords: Aspose.Cells | PresetShadowType | enum reflection C# | .NET | shape shadow presets | Enum.GetNames | Enum.GetValues | list enum members | dynamic configuration | Aspose.Cells.Drawing | C# code example | GitHub snippet
// Common Searches: list PresetShadowType enum Aspose.Cells | C# reflection get all shape shadow types | Aspose.Cells PresetShadowType values | how to enumerate PresetShadowType in .NET | Aspose.Cells shape shadow enumeration example
// Developer Intent: Retrieve a complete runtime list of PresetShadowType enum members and their numeric identifiers.
// Use Cases: Populate a dropdown or combo box with every shadow preset for user selection. | Validate that a user‑provided string corresponds to a valid PresetShadowType value. | Generate documentation or logs that display all supported shadow types for a workbook. | Automatically adapt code when new PresetShadowType members are added in future Aspose.Cells releases.
// AI Prompts: Generate C# code that returns a Dictionary<string,int> of PresetShadowType names and values using reflection. | Show how to bind the reflected PresetShadowType list to a WPF ComboBox. | Explain how to handle unknown PresetShadowType values when reading from external configuration files. | Write a PowerShell script that lists PresetShadowType members from the Aspose.Cells assembly.

using System;
using Aspose.Cells.Drawing;

// Shows how to obtain the PresetShadowType enumeration from Aspose.Cells.Drawing via reflection, retrieve each member name and its integer value with Enum.GetNames/GetValues, and print the list to the console. Useful for dynamic UI controls, validation, or documentation of shape shadow presets.
class Program
{
    static void Main()
    {
        // Obtain the enum type via reflection
        Type enumType = typeof(PresetShadowType);

        // Retrieve all names and corresponding values
        string[] names = Enum.GetNames(enumType);
        Array values = Enum.GetValues(enumType);

        Console.WriteLine("Available PresetShadowType members:");
        for (int i = 0; i < names.Length; i++)
        {
            Console.WriteLine($"{names[i]} = {(int)values.GetValue(i)}");
        }
    }
}
