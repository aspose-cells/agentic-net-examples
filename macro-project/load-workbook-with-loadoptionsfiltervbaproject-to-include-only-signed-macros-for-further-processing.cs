using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class LoadSignedVbaWorkbook
{
    static void Main()
    {
        try
        {
            const string inputPath = "SignedWorkbook.xlsm";
            const string outputPath = "ProcessedWorkbook.xlsm";

            // Verify that the input file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook. Aspose.Cells automatically detects the format.
            Workbook workbook = new Workbook(inputPath);

            // Check if a signed VBA project was loaded.
            if (workbook.VbaProject != null && workbook.VbaProject.IsSigned)
            {
                Console.WriteLine("VBA project is signed.");
                Console.WriteLine("Signature valid: " + workbook.VbaProject.IsValidSigned);
            }
            else
            {
                Console.WriteLine("No signed VBA project loaded.");
            }

            // Perform further processing on the workbook here...

            // Save the processed workbook.
            workbook.Save(outputPath, SaveFormat.Xlsm);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors.
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}