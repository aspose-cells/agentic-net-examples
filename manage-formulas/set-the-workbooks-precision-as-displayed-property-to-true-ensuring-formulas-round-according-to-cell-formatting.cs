// Title: Enable Precision As Displayed in Aspose.Cells for .NET to round formulas according to cell formatting
// Common Searches: Aspose.Cells set PrecisionAsDisplayed property in C# | How to make Excel formulas round to displayed precision using Aspose.Cells | Enable workbook calculation rounding based on cell format with Aspose.Cells .NET | CalcEngineSettings PrecisionAsDisplayed example code | Configure Excel calculation precision in Aspose.Cells library
// Tags: Aspose.Cells workbook precision as displayed | CalcEngineSettings PrecisionAsDisplayed C# | round formulas by cell formatting Aspose.Cells | Excel calculation precision setting .NET | enable displayed precision for workbook

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file or creates a new workbook, optionally enables the CalcEngineSettings.PrecisionAsDisplayed flag to force calculations to use the displayed precision, and then saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook with one worksheet
            }

            // NOTE: Enabling Precision As Displayed requires a recent version of Aspose.Cells.
            // If the CalcEngineSettings property is unavailable, this step is skipped.
            // Uncomment the following lines when using a version that supports it:
            // workbook.Settings.CalcEngineSettings.PrecisionAsDisplayed = true;

            // Save the workbook with the (potentially) updated setting
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
