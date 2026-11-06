// Title: Generate an Excel file with a WCAG contrast‑ratio matrix for the six theme accent colors using Aspose.Cells in C#
// Common Searches: how to calculate WCAG contrast ratio between theme accent colors and export to Excel with Aspose.Cells C# | C# Aspose.Cells example to create a contrast matrix for Excel theme accents | write color contrast values into an .xlsx file using Aspose.Cells in .NET | populate Excel header cells with background colors programmatically using Aspose.Cells | auto‑fit columns after inserting numeric data in an Aspose.Cells workbook
// Tags: Aspose.Cells create contrast ratio matrix | C# calculate WCAG contrast for Excel colors | populate Excel headers with background fills Aspose.Cells | auto-fit columns Aspose.Cells workbook | export contrast matrix to .xlsx using Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new workbook, defines six theme accent colors, writes their names as row and column headers with matching background fills, computes the WCAG contrast ratio for each color pair, inserts the rounded ratios into the matrix cells, auto‑fits the columns for readability, and saves the result as ContrastMatrix.xlsx. Helper methods calculate relative luminance and contrast ratios.
class ContrastMatrix
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Define theme accent colors (fallback if Theme class is unavailable)
            Color[] accents = new Color[]
            {
                Color.FromArgb(0, 112, 192),   // Accent1
                Color.FromArgb(255, 192, 0),   // Accent2
                Color.FromArgb(112, 173, 71),  // Accent3
                Color.FromArgb(255, 0, 0),     // Accent4
                Color.FromArgb(0, 176, 80),    // Accent5
                Color.FromArgb(255, 0, 255)    // Accent6
            };

            string[] accentNames = new string[]
            {
                "Accent1", "Accent2", "Accent3", "Accent4", "Accent5", "Accent6"
            };

            int startRow = 1; // Row index for the matrix header (0‑based)
            int startCol = 1; // Column index for the matrix header (0‑based)

            // Write column headers (accent names) and apply background colors
            for (int c = 0; c < accents.Length; c++)
            {
                Cell headerCell = cells[startRow, startCol + c + 1];
                headerCell.PutValue(accentNames[c]);

                Style style = workbook.CreateStyle();
                style.ForegroundColor = accents[c];
                style.Pattern = BackgroundType.Solid;
                headerCell.SetStyle(style);
            }

            // Write row headers (accent names) and apply background colors
            for (int r = 0; r < accents.Length; r++)
            {
                Cell headerCell = cells[startRow + r + 1, startCol];
                headerCell.PutValue(accentNames[r]);

                Style style = workbook.CreateStyle();
                style.ForegroundColor = accents[r];
                style.Pattern = BackgroundType.Solid;
                headerCell.SetStyle(style);
            }

            // Fill the matrix with contrast ratios
            for (int r = 0; r < accents.Length; r++)
            {
                for (int c = 0; c < accents.Length; c++)
                {
                    double ratio = ContrastRatio(accents[r], accents[c]);
                    cells[startRow + r + 1, startCol + c + 1].PutValue(Math.Round(ratio, 2));
                }
            }

            // Adjust column widths for readability
            sheet.AutoFitColumns();

            // Save the workbook
            workbook.Save("ContrastMatrix.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }

    // Calculates the WCAG contrast ratio between two colors
    static double ContrastRatio(Color c1, Color c2)
    {
        double l1 = RelativeLuminance(c1);
        double l2 = RelativeLuminance(c2);
        double lighter = l1 > l2 ? l1 : l2;
        double darker = l1 > l2 ? l2 : l1;
        return (lighter + 0.05) / (darker + 0.05);
    }

    // Computes the relative luminance of a color per WCAG definition
    static double RelativeLuminance(Color c)
    {
        double RsRGB = c.R / 255.0;
        double GsRGB = c.G / 255.0;
        double BsRGB = c.B / 255.0;

        double R = (RsRGB <= 0.03928) ? RsRGB / 12.92 : Math.Pow((RsRGB + 0.055) / 1.055, 2.4);
        double G = (GsRGB <= 0.03928) ? GsRGB / 12.92 : Math.Pow((GsRGB + 0.055) / 1.055, 2.4);
        double B = (BsRGB <= 0.03928) ? BsRGB / 12.92 : Math.Pow((BsRGB + 0.055) / 1.055, 2.4);

        return 0.2126 * R + 0.7152 * G + 0.0722 * B;
    }
}
