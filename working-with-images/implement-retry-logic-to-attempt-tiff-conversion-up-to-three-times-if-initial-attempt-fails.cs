// Title: Implement a three‑attempt retry loop for saving an Excel workbook as a TIFF image with Aspose.Cells in C#
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells and saves it as a TIFF, retrying the Save operation up to three times while logging each failure. | Add error handling that repeats the save operation up to three times before propagating the error.
// Common Searches: Aspose.Cells C# retry save as TIFF when exception occurs | How to add retry logic to workbook.Save with SaveFormat.Tiff in .NET | C# loop to attempt Excel to TIFF conversion multiple times using Aspose.Cells | Best practice for handling failed TIFF export with Aspose.Cells in C#
// Tags: Aspose.Cells retry save TIFF | C# workbook.Save exception handling | Excel to TIFF conversion retry loop | Aspose.Cells SaveFormat.Tiff error handling

using Aspose.Cells;
using System;

// // Loads an Excel workbook and attempts to save it as a TIFF file. If the save fails, the code retries up to three times, logs each error, and rethrows after the final unsuccessful attempt.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        string outputPath = "output.tiff";
        int maxAttempts = 3;
        int attempt = 0;
        bool success = false;

        while (attempt < maxAttempts && !success)
        {
            try
            {
                // Save the workbook as TIFF (using the provided save rule)
                workbook.Save(outputPath, SaveFormat.Tiff);
                success = true;
            }
            catch (Exception ex)
            {
                attempt++;
                if (attempt >= maxAttempts)
                {
                    // All attempts failed – rethrow or handle as needed
                    Console.WriteLine($"TIFF conversion failed after {attempt} attempts: {ex.Message}");
                    throw;
                }
                else
                {
                    // Log and retry
                    Console.WriteLine($"Attempt {attempt} failed: {ex.Message}. Retrying...");
                }
            }
        }

        if (success)
        {
            Console.WriteLine("TIFF conversion succeeded.");
        }
    }
}
