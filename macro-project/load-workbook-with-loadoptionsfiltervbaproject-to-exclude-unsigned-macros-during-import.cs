using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsm";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file not found – {inputPath}");
            return;
        }

        try
        {
            // Load the workbook. No explicit LoadOptions are required for .xlsm files.
            Workbook workbook = new Workbook(inputPath);

            // Display macro presence information.
            Console.WriteLine("HasMacro after load: " + workbook.HasMacro);

            // Save as a macro‑free XLSX file.
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors.
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}