// Title: Set Workbook Theme Font to Sans‑Serif (Arial) Globally with Aspose.Cells for .NET
// Description: Creates a new Workbook, accesses its DefaultStyle, and uses Font.SetName with FontSchemeType.Major and FontSchemeType.Minor to apply the Arial sans‑serif font to the entire workbook theme before saving.
// Keywords: Aspose.Cells | .NET | C# | Excel theme font | global font change | Arial | FontSchemeType | DefaultStyle | major font | minor font
// Common Searches: Aspose.Cells change workbook theme font to Arial | C# set global default font in Excel using Aspose.Cells | How to apply sans‑serif font to all cells with Aspose.Cells | Update major and minor font scheme in Aspose.Cells workbook
// Developer Intent: Apply a sans‑serif font to both major and minor theme schemes across the whole workbook.
// Use Cases: Generate brand‑compliant reports where the corporate sans‑serif font (e.g., Arial) is enforced for every cell. | Initialize new workbooks with a predefined theme font to maintain consistent styling in automated document pipelines.
// AI Prompts: Write C# code that changes the workbook theme font to Calibri for both major and minor schemes using Aspose.Cells and saves the file. | Explain the difference between FontSchemeType.Major and FontSchemeType.Minor in Aspose.Cells and how they affect Excel themes. | Create a reusable function that accepts any font name and applies it to the workbook's default style for both major and minor font schemes.

using System;
using Aspose.Cells;

namespace AsposeCellsThemeFontDemo
{
    // Creates a new Workbook, accesses its DefaultStyle, and uses Font.SetName with FontSchemeType.Major and FontSchemeType.Minor to apply the Arial sans‑serif font to the entire workbook theme before saving.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the default style which is applied globally to all cells
            Style defaultStyle = workbook.DefaultStyle;

            // Set a sans‑serif font (e.g., Arial) and assign it to the Major font scheme
            // The SetName method updates both the font name and its scheme type.
            defaultStyle.Font.SetName("Arial", FontSchemeType.Major);

            // Optionally, ensure the Minor scheme also uses a sans‑serif font
            defaultStyle.Font.SetName("Arial", FontSchemeType.Minor);

            // Save the workbook (lifecycle rule: save)
            workbook.Save("ThemeFontUpdated.xlsx");
        }
    }
}
