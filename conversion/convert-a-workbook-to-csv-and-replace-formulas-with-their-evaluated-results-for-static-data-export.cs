// Title: Convert an Excel workbook to CSV with formulas evaluated using Aspose.Cells for .NET
// Common Searches: how to save an .xlsx worksheet as a CSV file with calculated values using Aspose.Cells C# | Aspose.Cells convert Excel to CSV after evaluating all formulas | C# code to handle missing or password‑protected Excel file when exporting to CSV with Aspose.Cells | export first worksheet to CSV with formula results using Aspose.Cells .NET | calculate workbook formulas before CSV export Aspose.Cells example
// Tags: evaluate formulas before CSV export Aspose.Cells | save worksheet as CSV with calculated values C# | handle password protected workbook Aspose.Cells | fallback sample workbook creation Aspose.Cells | Aspose.Cells SaveFormat.Csv usage

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel file, creates a sample workbook if the source is missing or password‑protected, evaluates all formulas with CalculateFormula, and then saves the first worksheet as a CSV file containing the computed values, with comprehensive error handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.csv";

        try
        {
            Workbook workbook;

            if (File.Exists(inputPath))
            {
                // Attempt to load the workbook; handle password‑protected files gracefully
                try
                {
                    workbook = new Workbook(inputPath);
                }
                catch (Exception ex) when (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    Console.WriteLine("The input workbook is password protected. A sample workbook will be created instead.");
                    workbook = CreateSampleWorkbook();
                }
                catch (Exception ex)
                {
                    // Any other loading issue – create a sample workbook
                    Console.WriteLine($"Failed to load workbook: {ex.Message}. Creating a sample workbook.");
                    workbook = CreateSampleWorkbook();
                }
            }
            else
            {
                Console.WriteLine($"Input file '{inputPath}' not found. Creating a sample workbook.");
                workbook = CreateSampleWorkbook();
                try
                {
                    workbook.Save(inputPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save sample workbook: {ex.Message}");
                }
            }

            // Evaluate all formulas so cells contain calculated values
            workbook.CalculateFormula();

            // Export the first worksheet to CSV (formulas are saved as evaluated results)
            try
            {
                workbook.Save(outputPath, SaveFormat.Csv);
                Console.WriteLine($"CSV file saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save CSV file: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Creates a simple workbook with sample data and a formula
    private static Workbook CreateSampleWorkbook()
    {
        var wb = new Workbook();
        var ws = wb.Worksheets[0];
        ws.Cells["A1"].PutValue(10);
        ws.Cells["A2"].PutValue(20);
        ws.Cells["A3"].Formula = "=SUM(A1:A2)";
        return wb;
    }
}
