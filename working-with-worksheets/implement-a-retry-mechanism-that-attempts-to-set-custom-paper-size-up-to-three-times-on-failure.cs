// Title: Implement a three‑attempt retry loop for setting a custom paper size with Aspose.Cells in C#
// Common Searches: c# aspnet set worksheet page setup paper size with retry logic using Aspose.Cells | how to retry setting PaperSizeType in Aspose.Cells when an exception occurs | Aspose.Cells page setup custom paper size retry three times C# example | exception handling for setting worksheet paper size Aspose.Cells .NET | retry mechanism for pageSetup.PaperSize assignment Aspose.Cells
// Tags: retry loop for Aspose.Cells page setup | set worksheet paper size with error handling | Aspose.Cells custom paper size assignment | C# exception handling for page setup properties | Aspose.Cells workbook save after paper size retry

using Aspose.Cells;
using System;

// The sample creates a Workbook, accesses its first Worksheet, and uses a three‑attempt retry loop to assign PageSetup.PaperSize to PaperLetter. Each failure is logged, the final failure rethrows the exception, and the workbook is saved as CustomPaperSize.xlsx with overall exception handling.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Get the page setup object
            PageSetup pageSetup = sheet.PageSetup;

            // Retry mechanism: try up to three times to set paper size
            const int maxAttempts = 3;
            for (int attempt = 1; attempt <= maxAttempts; attempt++)
            {
                try
                {
                    // Use a predefined paper size that matches 8.5 x 11 inches
                    pageSetup.PaperSize = PaperSizeType.PaperLetter;

                    // If succeeded, exit loop
                    break;
                }
                catch (Exception ex)
                {
                    // If this was the last attempt, rethrow the exception
                    if (attempt == maxAttempts)
                    {
                        Console.WriteLine($"Failed to set paper size after {maxAttempts} attempts.");
                        throw;
                    }

                    // Log and continue to next attempt
                    Console.WriteLine($"Attempt {attempt} failed: {ex.Message}. Retrying...");
                }
            }

            // Save the workbook
            workbook.Save("CustomPaperSize.xlsx");
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            // General exception handling for unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
