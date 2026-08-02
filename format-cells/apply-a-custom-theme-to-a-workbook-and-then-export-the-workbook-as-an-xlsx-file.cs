// Title: C# – Apply a Custom Theme to an Aspose.Cells Workbook and Export as XLSX
// Description: Demonstrates how to create a new Workbook with Aspose.Cells for .NET, define a 12‑color custom theme, apply it using the CustomTheme method, and save the file as an XLSX document. Ideal for branding Excel reports or automating themed workbook generation.
// Keywords: Aspose.Cells | C# | .NET | CustomTheme | Excel theme colors | Workbook theme | XLSX export | Excel branding | automated Excel report | GitHub example | source code
// Common Searches: Aspose.Cells apply custom theme C# | How to set workbook theme colors with Aspose.Cells | Save workbook with custom theme as XLSX | CustomTheme method example Aspose.Cells .NET | Create branded Excel file using Aspose.Cells
// Developer Intent: Create a workbook, assign a custom color theme, and save it as an XLSX file.
// Use Cases: Generate corporate‑branded Excel reports with a predefined color palette. | Build reusable Excel templates that enforce consistent styling across automated outputs. | Produce documentation where hyperlink and visited‑link colors follow company style guidelines.
// AI Prompts: Show how to change the colors of an already applied custom theme in Aspose.Cells. | Provide code to load an existing workbook and replace its theme with a new custom theme using Aspose.Cells. | Explain how to read the current theme colors from a workbook in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsThemeDemo
{
    // Demonstrates how to create a new Workbook with Aspose.Cells for .NET, define a 12‑color custom theme, apply it using the CustomTheme method, and save the file as an XLSX document. Ideal for branding Excel reports or automating themed workbook generation.
    public class ApplyCustomTheme
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Define 12 colors for the custom theme
                Color[] customColors = new Color[]
                {
                    Color.FromArgb(255, 255, 255), // Background1
                    Color.FromArgb(0, 0, 0),       // Text1
                    Color.FromArgb(240, 240, 240), // Background2
                    Color.FromArgb(80, 80, 80),    // Text2
                    Color.FromArgb(0, 112, 192),   // Accent1
                    Color.FromArgb(255, 192, 0),   // Accent2
                    Color.FromArgb(112, 48, 160),  // Accent3
                    Color.FromArgb(0, 176, 80),    // Accent4
                    Color.FromArgb(255, 0, 0),     // Accent5
                    Color.FromArgb(0, 176, 240),   // Accent6
                    Color.FromArgb(0, 0, 255),     // Hyperlink
                    Color.FromArgb(128, 0, 128)    // Followed Hyperlink
                };

                // Apply the custom theme to the workbook
                workbook.CustomTheme("MyCustomTheme", customColors);

                // Save the workbook as XLSX
                workbook.Save("CustomThemeWorkbook.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error applying custom theme: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ApplyCustomTheme.Run();
        }
    }
}
