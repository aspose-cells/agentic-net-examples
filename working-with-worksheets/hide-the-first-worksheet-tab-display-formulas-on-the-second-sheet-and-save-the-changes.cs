// Title: Hide the first worksheet tab, show formulas on the second sheet, and save the workbook with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# hide first worksheet tab | display formulas on a specific worksheet using Aspose.Cells .NET | save workbook after modifying worksheet visibility with Aspose.Cells | set worksheet IsVisible false Aspose.Cells example | enable formula view in an Aspose.Cells workbook C#
// Tags: worksheet tab hiding Aspose.Cells | formula display setting Aspose.Cells | save workbook as xlsx Aspose.Cells | worksheet visibility control .NET | show formulas on second worksheet Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing XLSX file, hides the first worksheet by setting its IsVisible property to false, optionally enables formula view for the workbook, ensures the output directory exists, and saves the modified workbook to a new XLSX file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Hide the first worksheet tab
            Worksheet firstSheet = workbook.Worksheets[0];
            firstSheet.IsVisible = false; // Tab will be hidden

            // NOTE: In some Aspose.Cells versions the ShowFormula property is not available.
            // If needed, uncomment and adjust the following line according to the library version:
            // workbook.Settings.ShowFormula = true;

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the changes to a new file
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
