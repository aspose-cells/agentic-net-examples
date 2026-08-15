// Title: C# – Extract All Theme Colors from an XLSX Workbook with Aspose.Cells
// Description: Loads an XLSX file using Aspose.Cells, shows the workbook's theme name, iterates every ThemeColorType (except StyleColor) to output ARGB values, lists all colors actually used via CellsHelper.GetUsedColors, and saves the file to confirm the theme remains unchanged.
// Keywords: Aspose.Cells C# theme colors | GetThemeColor example | Workbook.Theme property | CellsHelper.GetUsedColors | extract Excel theme palette .NET | list workbook colors C# | Aspose.Cells theme extraction | Excel theme color definitions | C# Aspose.Cells sample code | theme color ARGB values
// Common Searches: Aspose.Cells retrieve all theme colors C# | how to list used colors in an Excel workbook with Aspose | C# code to get workbook theme name Aspose.Cells | save workbook after accessing theme colors Aspose | enumerate ThemeColorType values Aspose.Cells
// Developer Intent: Extract every theme color defined in an XLSX workbook and enumerate the colors actually used in the file.
// Use Cases: Generate a palette report to verify that a document follows corporate branding guidelines. | Create a color‑mapping tool that replaces theme colors with custom shades before distribution. | Audit multiple workbooks to ensure consistent theme usage across a product suite.
// AI Prompts: Write C# code with Aspose.Cells that loads an XLSX file and prints each ThemeColorType and its ARGB components, skipping StyleColor. | Provide a method that returns a Color[] containing all colors used in a workbook via CellsHelper.GetUsedColors. | Explain how to preserve the original workbook theme when saving after modifying cell styles with Aspose.Cells.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// Loads an XLSX file using Aspose.Cells, shows the workbook's theme name, iterates every ThemeColorType (except StyleColor) to output ARGB values, lists all colors actually used via CellsHelper.GetUsedColors, and saves the file to confirm the theme remains unchanged.
public class ThemeExtractor
{
    public static void Main()
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Path to the input workbook
        string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook from the file
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Display the workbook's theme name
        Console.WriteLine("Workbook theme: " + workbook.Theme);

        // Extract and display all theme color definitions
        Console.WriteLine("Theme colors:");
        foreach (ThemeColorType type in Enum.GetValues(typeof(ThemeColorType)))
        {
            // Skip internal StyleColor type
            if (type == ThemeColorType.StyleColor) continue;

            Color color = workbook.GetThemeColor(type);
            Console.WriteLine($"{type}: A={color.A}, R={color.R}, G={color.G}, B={color.B}");
        }

        // List all colors actually used in the workbook
        Color[] usedColors = CellsHelper.GetUsedColors(workbook);
        Console.WriteLine("Used colors in workbook:");
        foreach (Color c in usedColors)
        {
            Console.WriteLine($"A={c.A}, R={c.R}, G={c.G}, B={c.B}");
        }

        // Save the workbook to demonstrate that the theme persists
        try
        {
            workbook.Save("output_with_theme.xlsx");
            Console.WriteLine("Workbook saved as output_with_theme.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
