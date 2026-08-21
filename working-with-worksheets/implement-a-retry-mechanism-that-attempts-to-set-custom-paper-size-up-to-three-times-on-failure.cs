// Title: C# – Retry Setting Custom Paper Size on an Aspose.Cells Worksheet
// Description: Shows how to apply a 1.5 × 2.0 inches custom paper size to a worksheet using Aspose.Cells PageSetup.CustomPaperSize with a three‑attempt retry loop that logs each failure and rethrows after the final try, then saves the workbook.
// Keywords: Aspose.Cells | C# | custom paper size | PageSetup.CustomPaperSize | retry logic | exception handling | worksheet page setup | retry loop | transient errors | save workbook
// Common Searches: Aspose.Cells retry custom paper size | C# set custom paper size with retry | PageSetup.CustomPaperSize error handling | how to retry Aspose.Cells page setup | retry loop for worksheet page setup .NET
// Developer Intent: Add a retry mechanism that attempts to set a worksheet’s custom paper size up to three times and propagates the exception if every attempt fails.
// Use Cases: Ensuring page layout succeeds when printer drivers or temporary API glitches reject the dimensions. | Processing large batches of reports where intermittent exceptions may occur during page‑setup configuration. | Running automated document generation services that must guarantee successful page setup before saving.
// AI Prompts: Generate C# code that sets a custom paper size in Aspose.Cells with exponential back‑off and up to five retry attempts. | Refactor the retry loop to write each failure to a log file and use a configurable delay between attempts. | Explain which Aspose.Cells exceptions are transient for PageSetup.CustomPaperSize and how to retry only those cases.

using System;
using Aspose.Cells;

// Shows how to apply a 1.5 × 2.0 inches custom paper size to a worksheet using Aspose.Cells PageSetup.CustomPaperSize with a three‑attempt retry loop that logs each failure and rethrows after the final try, then saves the workbook.
public class SetCustomPaperSizeWithRetry
{
    public static void Main(string[] args)
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        PageSetup pageSetup = worksheet.PageSetup;

        // Desired custom paper size in inches
        double widthInches = 1.5;
        double heightInches = 2.0;

        // Retry parameters
        int maxRetry = 3;
        int attempt = 0;
        bool succeeded = false;

        // Attempt to set the custom paper size with retry logic
        while (attempt < maxRetry && !succeeded)
        {
            try
            {
                pageSetup.CustomPaperSize(widthInches, heightInches);
                succeeded = true; // Success, exit loop
            }
            catch (Exception ex)
            {
                attempt++;
                Console.WriteLine($"Attempt {attempt} failed: {ex.Message}");
                if (attempt >= maxRetry)
                {
                    Console.WriteLine("Maximum retry attempts reached. Rethrowing exception.");
                    throw; // Propagate the exception after final failure
                }
                // Optionally, add a short delay before retrying
                // System.Threading.Thread.Sleep(100);
            }
        }

        // Save the workbook using the standard save lifecycle
        workbook.Save("CustomPaperSizeRetryDemo.xlsx");
    }
}
