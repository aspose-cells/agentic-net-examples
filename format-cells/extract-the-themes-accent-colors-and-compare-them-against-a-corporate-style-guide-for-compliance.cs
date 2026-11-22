// Title: Extract Excel theme accent colors and verify corporate style guide compliance using Aspose.Cells for .NET (C#)
// Common Searches: C# code to read Excel theme accent colors with Aspose.Cells | how to compare workbook theme colors to a corporate color palette in .NET | validate Excel theme compliance against a style guide using Aspose.Cells | retrieve and list all six theme accent colors from an XLSX file in C# | check if Excel theme colors match predefined corporate colors programmatically
// Tags: Aspose.Cells extract theme accent colors | C# verify Excel theme against corporate palette | Excel theme compliance check using Aspose | retrieve workbook theme colors .NET | corporate style guide color validation Excel

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an XLSX workbook with Aspose.Cells, accesses its Theme via a dynamic reference, defines a list of corporate accent colors, iterates through the six theme accent slots, retrieves each accent color, compares it to the corporate palette for exact matches, and prints whether each accent is COMPLIANT or NON‑COMPLIANT.
class ThemeComplianceChecker
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: File not found – {inputPath}");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        try
        {
            // Access the workbook's theme using dynamic to avoid compile‑time dependency on Theme class
            dynamic theme = workbook.Theme;

            // Define corporate accent colors (replace with actual values as needed)
            List<Color> corporateAccents = new List<Color>
            {
                Color.FromArgb(255, 0, 120, 215),   // Corporate Accent 1
                Color.FromArgb(255, 232, 17, 35),   // Corporate Accent 2
                Color.FromArgb(255, 255, 185, 0),   // Corporate Accent 3
                Color.FromArgb(255, 0, 153, 0),     // Corporate Accent 4
                Color.FromArgb(255, 112, 48, 160),  // Corporate Accent 5
                Color.FromArgb(255, 255, 102, 0)    // Corporate Accent 6
            };

            // Check each of the six theme accent colors
            for (int i = 0; i < 6; i++)
            {
                // Retrieve the accent color from the theme (index 0‑5)
                Color accentColor = theme.GetAccentColor(i);

                // Determine compliance by checking if the accent matches any corporate color
                bool isCompliant = corporateAccents.Exists(c => ColorsAreEqual(c, accentColor));

                // Output the result
                Console.WriteLine($"Accent {i + 1}: #{accentColor.ToArgb():X8} " +
                                  (isCompliant ? "COMPLIANT" : "NON‑COMPLIANT"));
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Runtime error: {ex.Message}");
        }
    }

    // Helper method to compare two Color objects for exact equality
    private static bool ColorsAreEqual(Color a, Color b)
    {
        return a.ToArgb() == b.ToArgb();
    }
}
