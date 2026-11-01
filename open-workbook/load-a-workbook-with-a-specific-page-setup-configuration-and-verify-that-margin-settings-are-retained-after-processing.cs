// Title: Load an Excel workbook, set custom page margins, save it, and confirm margin values are preserved with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells set worksheet page margins and verify after saving | How to check if Excel margin settings are retained after reloading with Aspose.Cells | Aspose.Cells .NET example for modifying and persisting page setup margins | Validate custom left, right, top, bottom margins in a saved Excel file using Aspose.Cells | Load, edit, and re‑load Excel workbook margins programmatically in C#
// Tags: Aspose.Cells worksheet page margins | C# set custom page margins Excel | validate margin retention after save | load and modify Excel page setup Aspose | reloaded workbook margin verification

using System;
using System.IO;
using Aspose.Cells;

// This C# example uses Aspose.Cells to ensure an Excel file exists, loads the workbook, sets left, right, top, and bottom margins via the worksheet's PageSetup, saves the file, reloads it, and compares the reloaded margin values to the originals, outputting whether the custom margin settings were retained.
class Program
{
    static void Main()
    {
        // Paths for input and output workbooks
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Ensure the input workbook exists; create a blank one if it does not
            if (!File.Exists(inputPath))
            {
                Workbook newWb = new Workbook();
                newWb.Worksheets[0].Name = "Sheet1";
                newWb.Save(inputPath);
            }

            // Load the workbook
            Workbook wb = new Workbook(inputPath);

            // Access the first worksheet's page setup
            PageSetup pageSetup = wb.Worksheets[0].PageSetup;

            // Set custom margin values (in inches)
            pageSetup.LeftMargin = 0.5;
            pageSetup.RightMargin = 0.5;
            pageSetup.TopMargin = 0.75;
            pageSetup.BottomMargin = 0.75;

            // Save the workbook after setting margins
            wb.Save(outputPath);

            // Reload the saved workbook to verify margins are retained
            Workbook wbReloaded = new Workbook(outputPath);
            PageSetup reloadedSetup = wbReloaded.Worksheets[0].PageSetup;

            // Verify each margin
            bool marginsMatch =
                Math.Abs(reloadedSetup.LeftMargin - 0.5) < 0.0001 &&
                Math.Abs(reloadedSetup.RightMargin - 0.5) < 0.0001 &&
                Math.Abs(reloadedSetup.TopMargin - 0.75) < 0.0001 &&
                Math.Abs(reloadedSetup.BottomMargin - 0.75) < 0.0001;

            // Output verification result
            Console.WriteLine(marginsMatch
                ? "Margin settings are retained after processing."
                : "Margin settings were not retained.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
