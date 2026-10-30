// Title: Check for existing workbook theme before applying bulk theme changes with Aspose.Cells for .NET
// Common Searches: how to verify if an Excel workbook has a theme using Aspose.Cells C# | batch process Excel files to detect theme presence with Aspose.Cells | skip theme modification for workbooks without a theme Aspose.Cells .NET | list .xlsx files and check Theme property in Aspose.Cells | determine if workbook.Theme is set before updating colors Aspose.Cells
// Tags: workbook theme detection Aspose.Cells | bulk workbook processing C# Aspose.Cells | theme property check Excel .NET | iterate .xlsx files Aspose.Cells | conditional theme update Aspose.Cells

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

// The example iterates over all .xlsx files in a folder, loads each workbook with Aspose.Cells for .NET, checks whether the workbook's Theme property is non‑empty to determine if a theme exists, logs the result, and saves the workbook (unchanged) to an output directory.
class Program
{
    static void Main()
    {
        // Folder containing the workbooks to process
        string inputFolder = @"C:\InputWorkbooks";
        // Folder where the updated workbooks will be saved
        string outputFolder = @"C:\OutputWorkbooks";

        // Ensure the input directory exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder does not exist: {inputFolder}");
            return;
        }

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Process each .xlsx file in the input folder
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx"))
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                // Load the workbook (lifecycle rule: load)
                Workbook workbook = new Workbook(filePath);

                // Determine whether the workbook has a theme defined
                bool hasTheme = !string.IsNullOrEmpty(workbook.Theme);
                Console.WriteLine($"{Path.GetFileName(filePath)} - Contains Theme: {hasTheme}");

                // If a theme exists, you could modify it here using the appropriate API.
                // The ThemeColorScheme property is not available in this version of Aspose.Cells,
                // so theme modification is omitted.

                // Save the (potentially) modified workbook (lifecycle rule: save)
                string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
                workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }
    }
}
