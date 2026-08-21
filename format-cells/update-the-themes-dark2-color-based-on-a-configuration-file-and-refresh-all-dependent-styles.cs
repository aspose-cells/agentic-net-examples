// Title: C# Example: Update Excel Dark2 (Background2) Theme Color from a Config File using Aspose.Cells
// Description: Loads an existing workbook, reads a hex color from a text configuration file, converts it to a System.Drawing.Color, applies the value to the Dark2 (Background2) theme via Workbook.SetThemeColor, and saves the updated file. Includes basic error handling for missing files and invalid color strings.
// Keywords: Aspose.Cells C# | SetThemeColor | ThemeColorType.Background2 | Dark2 theme color | Excel theme update | hex color from config | read color file C# | Workbook.SetThemeColor example | .NET Excel styling | GitHub Aspose.Cells sample
// Common Searches: how to change Dark2 theme color in Excel with Aspose.Cells .NET | C# read hex color from file and set Excel theme | Aspose.Cells SetThemeColor Background2 example | update Excel theme colors from configuration file | apply corporate brand color to Excel workbook using Aspose
// Developer Intent: Programmatically set the Dark2 (Background2) theme color of an Excel workbook based on a hex value stored in an external configuration file.
// Use Cases: Enforce corporate branding by loading a brand color from a central config and applying it to all generated reports. | Allow end‑users to customize report appearance by selecting a color saved in a simple text file. | Batch‑process multiple workbooks, updating their theme colors from individual config files to maintain visual consistency.
// AI Prompts: Generate C# code that reads a hex color from a JSON configuration file and updates the Dark2 (Background2) theme color in an Aspose.Cells workbook. | Show how to add comprehensive validation for hex strings and file existence when using Workbook.SetThemeColor for Background2. | Explain how to refresh or reapply cell styles after changing the Dark2 theme color with Aspose.Cells.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// Loads an existing workbook, reads a hex color from a text configuration file, converts it to a System.Drawing.Color, applies the value to the Dark2 (Background2) theme via Workbook.SetThemeColor, and saves the updated file. Includes basic error handling for missing files and invalid color strings.
class UpdateDark2Theme
{
    static void Main()
    {
        try
        {
            // Paths for the workbook and configuration file
            string workbookPath = "input.xlsx";
            string configPath = "themeconfig.txt";   // Expected format: #RRGGBB or RRGGBB

            // Verify that the required files exist
            if (!File.Exists(workbookPath))
                throw new FileNotFoundException($"Workbook file not found: {workbookPath}");
            if (!File.Exists(configPath))
                throw new FileNotFoundException($"Configuration file not found: {configPath}");

            // Load the existing workbook
            Workbook workbook = new Workbook(workbookPath);

            // Read the color value from the configuration file
            string colorString = File.ReadAllText(configPath).Trim();

            // Ensure the color string starts with '#'
            if (!colorString.StartsWith("#"))
                colorString = "#" + colorString;

            // Convert the string to a System.Drawing.Color
            Color dark2Color = ColorTranslator.FromHtml(colorString);

            // Update the Dark2 (Background2) theme color
            workbook.SetThemeColor(ThemeColorType.Background2, dark2Color);

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
