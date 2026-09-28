// Title: Use a three‑attempt retry loop to set a custom paper size on an Aspose.Cells worksheet in C#
// AI Prompts: Write C# code that creates an Aspose.Cells workbook, accesses the first worksheet, and assigns a specific PaperSizeType to PageSetup inside a retry loop that tries up to three times before rethrowing the exception. | Generate a .NET example showing how to wrap sheet.PageSetup.PaperSize assignment with exception handling, a counter, and a maximum of three retries, then save the workbook.
// Common Searches: Aspose.Cells C# retry setting worksheet page setup paper size on failure | how to implement retry logic for PaperSizeType assignment in Aspose.Cells | C# example with multiple attempts to set custom paper size in Excel workbook using Aspose | exception handling for worksheet PageSetup PaperSize in Aspose.Cells .NET | retry loop for setting A4 paper size with Aspose.Cells workbook
// Tags: retry loop Aspose.Cells PageSetup PaperSize | exception handling worksheet paper size Aspose.Cells | C# Aspose.Cells set custom paper size with retries | three attempts paper size assignment Aspose.Cells | Aspose.Cells workbook save after page setup retry

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new workbook, accesses the first worksheet, and attempts to set the worksheet's PageSetup PaperSize to A4 inside a retry loop that makes up to three attempts. If all attempts fail, the exception is rethrown. Finally, the workbook is saved as CustomPaperSizeWorkbook.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Retry mechanism: try to set a paper size up to three times
            const int maxAttempts = 3;
            int attempt = 0;
            bool isSet = false;

            while (attempt < maxAttempts && !isSet)
            {
                try
                {
                    // Set a predefined paper size (A4) – use the correct enum value
                    sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;

                    // If no exception, the operation succeeded
                    isSet = true;
                }
                catch (Exception)
                {
                    attempt++;

                    // If maximum attempts reached, rethrow the exception
                    if (attempt >= maxAttempts)
                    {
                        throw;
                    }
                }
            }

            // Define output file path
            string outputPath = "CustomPaperSizeWorkbook.xlsx";

            // Save the workbook (lifecycle save rule)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            // Log or handle unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
