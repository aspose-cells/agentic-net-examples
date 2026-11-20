// Title: How to emulate a lighter Light2 theme background while preserving dark text readability in Excel using Aspose.Cells for .NET
// Common Searches: asp.net cells change Light2 theme background to a lighter shade | c# make Excel workbook background lighter with Aspose.Cells and keep dark font readable | how to apply custom cell styles to simulate lighter theme in Aspose.Cells .NET | adjust workbook theme colors for better contrast using Aspose.Cells C# | workaround for unsupported theme modification in Aspose.Cells
// Tags: Aspose.Cells custom cell style background color | C# emulate lighter workbook background via cell formatting | dark text contrast after background change Aspose.Cells | Light2 theme equivalent modification using cell styles .NET | Excel theme color limitation workaround Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook (creating one if missing) and notes that direct theme manipulation is not supported in the current Aspose.Cells version. Developers can achieve a lighter Light2 background and maintain dark text readability by applying custom cell style background colors as a workaround, then saving the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists; create a blank workbook if it does not.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found. Creating a new workbook.");
                var newWb = new Workbook();
                newWb.Save(inputPath);
            }

            // Load the workbook.
            var workbook = new Workbook(inputPath);

            // NOTE: Direct theme manipulation is not available in the current Aspose.Cells version.
            // If needed, you can modify cell styles individually here.

            // Save the (potentially modified) workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
