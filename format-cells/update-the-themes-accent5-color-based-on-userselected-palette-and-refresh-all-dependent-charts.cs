using System;
using System.Drawing;
using Aspose.Cells;

class UpdateAccent5Theme
{
    static void Main()
    {
        // Path to the workbook that contains charts dependent on the theme
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Load the workbook
        Workbook workbook = new Workbook(inputPath);

        // Example of a user‑selected color (replace with actual selection logic)
        Color userSelectedColor = Color.FromArgb(255, 165, 0); // Orange

        // Update the Accent5 theme color
        workbook.SetThemeColor(ThemeColorType.Accent5, userSelectedColor);

        // Optional: verify the change
        Color verified = workbook.GetThemeColor(ThemeColorType.Accent5);
        Console.WriteLine($"Accent5 theme color set to: {verified}");

        // Save the workbook; charts that use Accent5 will reflect the new color
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}