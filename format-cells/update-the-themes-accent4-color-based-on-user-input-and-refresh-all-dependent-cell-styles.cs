// Title: Change the Accent4 color of an Excel workbook’s theme from a hex string and refresh dependent styles using Aspose.Cells for .NET
// Common Searches: how to set the Accent4 theme color in an existing .xlsx file using Aspose.Cells C# | update Excel theme accent color from user‑entered hex value with Aspose.Cells | refresh cell style formatting after modifying workbook theme in C# | apply custom hex color to Excel theme Accent4 and recalculate formulas using Aspose.Cells | prompt user for color and change Excel theme accent programmatically in .NET
// Tags: set theme accent4 color Aspose.Cells | apply hex color to Excel theme C# | recalculate workbook formulas after theme change | dynamic theme manipulation Aspose.Cells | refresh dependent cell styles Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example prompts the user for a hex‑encoded color, converts it to a System.Drawing.Color, loads an existing workbook, uses a dynamic reference to modify the workbook’s Theme Accent4 color, recalculates all formulas to propagate style changes, and saves the updated file.
class UpdateThemeAccent4
{
    static void Main()
    {
        // Prompt user for the new Accent4 color in hex format (e.g., #FF5733)
        Console.Write("Enter new Accent4 color (hex, e.g., #FF5733): ");
        string hexColor = Console.ReadLine();

        if (string.IsNullOrWhiteSpace(hexColor))
        {
            Console.WriteLine("No color entered.");
            return;
        }

        // Convert the hex string to a System.Drawing.Color
        Color newAccent4;
        try
        {
            newAccent4 = ColorTranslator.FromHtml(hexColor);
        }
        catch (Exception)
        {
            Console.WriteLine("Invalid color format.");
            return;
        }

        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Ensure the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access and modify the theme's Accent4 color using dynamic to avoid compile‑time dependency on Theme type
            dynamic theme = workbook.Theme;
            theme.ColorScheme.Accent4 = newAccent4;

            // Recalculate formulas to reflect any dependent changes
            workbook.CalculateFormula();

            // Save the modified workbook
            workbook.Save(outputPath);

            Console.WriteLine("Accent4 color updated and workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
