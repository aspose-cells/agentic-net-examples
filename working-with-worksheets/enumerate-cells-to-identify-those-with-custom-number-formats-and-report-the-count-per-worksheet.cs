// Title: C# – Count Cells with Custom Number Formats per Worksheet using Aspose.Cells
// Description: Loads an Excel file with Aspose.Cells, iterates each worksheet, enumerates instantiated cells, checks the Style.Custom property, counts cells that have a custom number format, outputs the count per sheet, and saves the workbook unchanged.
// Keywords: Aspose.Cells | C# | custom number format | count cells | worksheet enumeration | Style.Custom | Excel audit | cell format detection | .NET | enumerate cells
// Common Searches: Aspose.Cells count cells with custom format | how to detect custom number formats in Excel using C# | enumerate cells and check Style.Custom Aspose.Cells | C# get number of custom formatted cells per worksheet | Aspose.Cells custom number format audit
// Developer Intent: Count the number of cells that use a custom number format on each worksheet of an Excel workbook.
// Use Cases: Audit workbooks for unintended custom formats | Generate a summary report of custom formatted cells per sheet | Validate formatting compliance before publishing | Create a quality‑control dashboard showing format usage | Identify sheets that may need format standardization
// AI Prompts: Generate C# code with Aspose.Cells that lists the addresses of cells containing custom number formats and writes the results to a CSV file. | Show how to extend the sample to group counts by the custom format string. | Explain how to run the example in a .NET Core console app and handle large workbooks efficiently. | Provide a PowerShell script that calls the compiled program and captures its output. | Create a unit test that verifies the custom format counter returns expected values for a sample workbook.

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel file with Aspose.Cells, iterates each worksheet, enumerates instantiated cells, checks the Style.Custom property, counts cells that have a custom number format, outputs the count per sheet, and saves the workbook unchanged.
public class CustomNumberFormatCounter
{
    public static void Run()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                long customFormatCount = 0;
                Cells cells = sheet.Cells;

                // Enumerate all instantiated cells
                foreach (Cell cell in cells)
                {
                    Style style = cell.GetStyle();

                    // If the Custom property is not empty, the cell uses a custom number format
                    if (!string.IsNullOrEmpty(style.Custom))
                    {
                        customFormatCount++;
                    }
                }

                // Report the count for the current worksheet
                Console.WriteLine($"Worksheet \"{sheet.Name}\" contains {customFormatCount} cells with custom number formats.");
            }

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (no modifications made)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Entry point for the application
    public static void Main()
    {
        Run();
    }
}
