// Title: Change the default workbook font color to black for better dark‑text contrast using Aspose.Cells for .NET
// Common Searches: how to set default font color to black in an Excel workbook with Aspose.Cells C# | Aspose.Cells change default style font color for dark text readability | C# code to update default workbook style font color using Aspose.Cells | make dark text readable by setting default font color in Aspose.Cells .NET | Aspose.Cells default style font color black example
// Tags: default style font color black Aspose.Cells | workbook default style update C# | dark text contrast Excel Aspose.Cells | modify workbook default style Aspose.Cells | set default font black .NET

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// The example loads an existing .xlsx file with Aspose.Cells, changes the workbook’s default style font color to black to improve dark‑text contrast, and saves the modified workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // NOTE: Theme color modification is omitted because the required enum value may not be available
            // in older Aspose.Cells versions. The default style is updated instead.

            // Ensure readability of dark text by setting the default font color to black
            Style defaultStyle = workbook.DefaultStyle;
            defaultStyle.Font.Color = Color.Black;
            workbook.DefaultStyle = defaultStyle;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
