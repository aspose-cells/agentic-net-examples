// Title: Add three‑attempt retry logic to Excel‑to‑TIFF conversion using Aspose.Cells in C#
// Common Searches: c# Aspose.Cells retry saving workbook as TIFF when an error occurs | how to implement a retry loop for Excel to TIFF conversion with Aspose.Cells | Aspose.Cells export Excel file to TIFF with multiple attempts in .NET | handle transient failures during workbook.Save to TIFF in C# | retry mechanism for SaveFormat.Tiff in Aspose.Cells applications
// Tags: Aspose.Cells retry logic for TIFF export | C# workbook.Save error handling | Excel to TIFF conversion with retry | SaveFormat.Tiff retry implementation | transient failure handling in Aspose.Cells

using System;
using Aspose.Cells;

// Demonstrates how to wrap the Aspose.Cells workbook.Save call in a three‑attempt retry loop, loading an Excel file and exporting it to TIFF while handling exceptions, pausing between attempts, and rethrowing a detailed error if all retries fail.
class Program
{
    static void Main()
    {
        string inputFile = "input.xlsx";
        string outputFile = "output.tiff";

        try
        {
            ConvertToTiff(inputFile, outputFile);
            Console.WriteLine("Conversion succeeded.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Conversion failed after retries: {ex.Message}");
        }
    }

    static void ConvertToTiff(string inputPath, string outputPath)
    {
        const int maxAttempts = 3;
        int attempt = 0;
        bool success = false;
        Exception lastException = null;

        while (attempt < maxAttempts && !success)
        {
            attempt++;
            try
            {
                // Load the workbook (lifecycle rule: load)
                Workbook workbook = new Workbook(inputPath);

                // Save as TIFF (lifecycle rule: save)
                workbook.Save(outputPath, SaveFormat.Tiff);

                success = true; // conversion succeeded
            }
            catch (Exception ex)
            {
                lastException = ex;
                Console.WriteLine($"Attempt {attempt} failed: {ex.Message}");
                // Optionally add a short delay before retrying
                System.Threading.Thread.Sleep(500);
            }
        }

        if (!success && lastException != null)
        {
            // Propagate the last exception after all retries have been exhausted
            throw new InvalidOperationException($"Failed to convert '{inputPath}' to TIFF after {maxAttempts} attempts.", lastException);
        }
    }
}
