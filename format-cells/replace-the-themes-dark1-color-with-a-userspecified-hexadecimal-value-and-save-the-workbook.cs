// Title: Set Dark1 Theme Color via Hex Code in an Excel Workbook with Aspose.Cells for .NET (C#)
// Description: C# sample that creates a new workbook, converts a user‑provided hexadecimal string (e.g., "#FF5733") to a System.Drawing.Color, assigns it to the Dark1 (Background1) theme element using Workbook.SetThemeColor, and saves the file to a specified path.
// Keywords: Aspose.Cells | C# | SetThemeColor | Dark1 theme color | hexadecimal color | Excel workbook styling | custom theme color | Background1 | ColorTranslator | programmatic Excel formatting
// Common Searches: Aspose.Cells change Dark1 theme color | C# set Excel theme color hex | How to modify workbook theme color with Aspose.Cells | Set custom theme color in Excel using .NET | Replace default Dark1 color Aspose.Cells
// Developer Intent: Replace the workbook’s Dark1 (Background1) theme color with a hex value supplied by the user and save the updated workbook.
// Use Cases: Brand a generated report by matching the Dark1 theme color to a corporate hex code. | Produce department‑specific workbooks, each using a different accent color for the Dark1 element. | Enable end‑users to pick a color in a UI and export an Excel file whose theme reflects that selection.
// AI Prompts: Write a C# method that validates a hex string, updates the Dark1 theme color of an existing workbook with Aspose.Cells, and saves the file. | Add robust error handling for hex‑to‑Color conversion and SetThemeColor calls in the provided example. | Show how to read the current Dark1 color from a workbook before applying a new hex value using Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsThemeDemo
{
    // C# sample that creates a new workbook, converts a user‑provided hexadecimal string (e.g., "#FF5733") to a System.Drawing.Color, assigns it to the Dark1 (Background1) theme element using Workbook.SetThemeColor, and saves the file to a specified path.
    public class ReplaceDark1Color
    {
        /// <param name="hexColor">Hexadecimal color string (e.g., "#FF5733").</param>
        /// <param name="outputPath">Path where the workbook will be saved.</param>
        public static void Run(string hexColor, string outputPath)
        {
            // Create a new workbook (default theme is applied)
            Workbook workbook = new Workbook();

            // Convert the hexadecimal string to a System.Drawing.Color.
            // ColorTranslator supports strings with or without the leading '#'.
            Color newColor = ColorTranslator.FromHtml(hexColor);

            // Set the theme color for Dark1 (Background1) to the new color.
            workbook.SetThemeColor(ThemeColorType.Background1, newColor);

            // Save the workbook to the specified location.
            workbook.Save(outputPath);
        }

        // Example usage
        public static void Main()
        {
            // User‑specified hexadecimal color.
            string userHex = "#4A90E2";

            // Desired output file.
            string filePath = "WorkbookWithCustomDark1.xlsx";

            Run(userHex, filePath);

            Console.WriteLine($"Workbook saved with Dark1 color set to {userHex} at '{filePath}'.");
        }
    }
}
