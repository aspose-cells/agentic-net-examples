using System;
using Aspose.Cells;

namespace AsposeCellsRetryDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Desired custom paper size in inches
            double widthInches = 1.5;
            double heightInches = 1.5;

            // Maximum number of attempts to set the custom paper size
            int maxAttempts = 3;
            int attempt = 0;
            bool isSet = false;

            // Retry loop
            while (attempt < maxAttempts && !isSet)
            {
                try
                {
                    // Attempt to set the custom paper size
                    worksheet.PageSetup.CustomPaperSize(widthInches, heightInches);
                    isSet = true; // Success, exit loop
                }
                catch (Exception ex)
                {
                    attempt++;
                    // If we've exhausted all attempts, rethrow the exception
                    if (attempt >= maxAttempts)
                    {
                        Console.WriteLine($"Failed to set custom paper size after {maxAttempts} attempts: {ex.Message}");
                        throw;
                    }
                    // Optionally, log the failure and continue retrying
                    Console.WriteLine($"Attempt {attempt} failed: {ex.Message}. Retrying...");
                }
            }

            // Save the workbook to demonstrate the applied setting
            workbook.Save("CustomPaperSizeWithRetry.xlsx");
            Console.WriteLine("Workbook saved successfully.");
        }
    }
}