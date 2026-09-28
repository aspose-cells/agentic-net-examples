// Title: Set Aspose.Cells CalculationOptions.ThreadCount to 4 for controlled formula evaluation in C#
// AI Prompts: Generate C# code that configures Aspose.Cells CalculationOptions.ThreadCount = 4 before calling Workbook.CalculateFormula(). | Show an example of limiting the number of parallel calculation threads to four when evaluating formulas with Aspose.Cells in a .NET application. | Explain how to adjust Aspose.Cells calculation thread pool to a fixed size of four threads and then save the workbook.
// Common Searches: how to restrict Aspose.Cells formula calculation to a specific number of threads in C# | Aspose.Cells CalculationOptions ThreadCount property usage example | limit parallel formula evaluation in Aspose.Cells .NET | set maximum calculation threads for Excel workbook using Aspose.Cells | C# Aspose.Cells control calculation thread pool size
// Tags: Aspose.Cells calculation thread count | C# set CalculationOptions.ThreadCount | limit parallel formula evaluation Aspose.Cells | configure Aspose.Cells calculation options | control workbook.CalculateFormula threads

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates loading or creating a workbook, adding data and formulas, configuring Aspose.Cells CalculationOptions.ThreadCount to 4 to restrict parallel formula evaluation, calculating the formulas, and saving the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Load existing workbook if present; otherwise create a new blank workbook
            string inputPath = "input.xlsx";
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
            }

            // Add sample data and formulas
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;
            cells["A1"].PutValue(10);
            cells["A2"].Formula = "=A1*2";
            cells["A3"].Formula = "=SUM(A1:A2)";

            // Calculate all formulas (default calculation options)
            try
            {
                workbook.CalculateFormula();
            }
            catch (Exception calcEx)
            {
                Console.WriteLine($"Formula calculation error: {calcEx.Message}");
            }

            // Ensure the output directory exists
            string outputPath = "output.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
