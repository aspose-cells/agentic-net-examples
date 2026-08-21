// Title: C# – Change Dark1 (Background1) Theme Color with a Hex Code using Aspose.Cells and Save the Workbook
// Description: Demonstrates how to convert a user‑provided hex string (e.g., "#FF5733") to a System.Drawing.Color, apply it to the Dark1 (Background1) theme color via Workbook.SetThemeColor, and save the workbook as "Workbook_With_Custom_Dark1.xlsx".
// Keywords: Aspose.Cells SetThemeColor | C# change Dark1 theme color | Background1 hex color Aspose | custom theme color Excel .NET | Aspose.Cells theme color example
// Common Searches: how to set Dark1 theme color Aspose.Cells C# | replace Background1 color with hex value in Excel using Aspose | Aspose.Cells change theme color programmatically | save workbook after modifying theme colors Aspose.Cells
// Developer Intent: Apply a user‑specified hex color to the workbook’s Dark1 (Background1) theme and persist the change.
// Use Cases: Branding: enforce corporate brand color across all generated reports. | Dynamic styling: let end‑users pick a theme color when creating spreadsheets. | Template preparation: build a starter workbook with a custom Dark1 color for downstream automation.
// AI Prompts: Write C# code that reads a hex color from input, sets the Dark1 (Background1) theme color in an Aspose.Cells workbook, and saves the file. | Explain the role of ThemeColorType.Background1 in Aspose.Cells and how to convert a hex string to System.Drawing.Color. | Show how to load a hex value from a JSON configuration, apply it to the Dark1 theme, and export the workbook with Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;

// Demonstrates how to convert a user‑provided hex string (e.g., "#FF5733") to a System.Drawing.Color, apply it to the Dark1 (Background1) theme color via Workbook.SetThemeColor, and save the workbook as "Workbook_With_Custom_Dark1.xlsx".
class ReplaceDark1ThemeColor
{
    static void Main()
    {
        // User‑specified hexadecimal color (e.g., "#FF5733")
        string hexColor = "#FF5733";

        // Convert the hex string to a System.Drawing.Color
        Color newColor = ColorTranslator.FromHtml(hexColor);

        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();

        // Replace the Dark1 (Background1) theme color with the new color
        workbook.SetThemeColor(ThemeColorType.Background1, newColor);

        // Save the workbook (lifecycle rule: save)
        workbook.Save("Workbook_With_Custom_Dark1.xlsx");
    }
}
