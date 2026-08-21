// Title: C# – Change Hyperlink Theme Color to DodgerBlue and Save Workbook with Aspose.Cells
// Description: Creates a new workbook, sets the Hyperlink theme color to DodgerBlue using SetThemeColor, adds a sample hyperlink, and saves the file as HyperlinkThemeUpdated.xlsx.
// Keywords: Aspose.Cells SetThemeColor hyperlink | C# change hyperlink color Excel | DodgerBlue theme color Aspose | update hyperlink theme color .NET | save workbook after theme change | Aspose.Cells hyperlink example | Excel theme color customization C#
// Common Searches: How to set hyperlink theme color in Excel using Aspose.Cells C# | Aspose.Cells SetThemeColor method example | Change hyperlink color to custom blue in .NET | Save workbook after modifying theme colors Aspose | Add hyperlink with custom theme color using Aspose.Cells
// Developer Intent: Apply a custom blue to the Hyperlink theme and persist the workbook.
// Use Cases: Brand‑consistent reports where all hyperlinks match corporate blue guidelines. | Templates that automatically apply a UI‑matching hyperlink shade for better visual integration. | Automated Excel generation with clickable links that use a specific blue for accessibility and style.
// AI Prompts: Show a C# snippet that changes the Hyperlink theme color to a custom RGB value with Aspose.Cells and saves the workbook. | Provide an example that updates the Hyperlink theme color, inserts multiple hyperlinks, and writes the file using Aspose.Cells for .NET. | Explain how SetThemeColor influences existing hyperlinks and whether a refresh is needed after changing the theme.

using System;
using System.Drawing;
using Aspose.Cells;

namespace UpdateHyperlinkThemeColor
{
    // Creates a new workbook, sets the Hyperlink theme color to DodgerBlue using SetThemeColor, adds a sample hyperlink, and saves the file as HyperlinkThemeUpdated.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Define a new shade of blue for the Hyperlink theme color
                Color newBlue = Color.FromArgb(30, 144, 255); // DodgerBlue

                // Update the Hyperlink theme color
                workbook.SetThemeColor(ThemeColorType.Hyperlink, newBlue);

                // Add a sample hyperlink to demonstrate the theme color
                Worksheet sheet = workbook.Worksheets[0];
                Cell cell = sheet.Cells["A1"];
                cell.PutValue("Aspose.Cells");

                // Use the Hyperlinks collection to add an external hyperlink
                sheet.Hyperlinks.Add("A1", 0, 0, "https://www.aspose.com");

                // Save the workbook
                workbook.Save("HyperlinkThemeUpdated.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
