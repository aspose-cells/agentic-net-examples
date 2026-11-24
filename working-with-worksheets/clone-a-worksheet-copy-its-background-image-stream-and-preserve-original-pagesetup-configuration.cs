// Title: Clone an Excel worksheet with its background image and keep original page‑setup settings using Aspose.Cells for .NET
// Common Searches: aspnet clone worksheet preserving background image Aspose.Cells | copy worksheet page setup and background image with Aspose.Cells C# | how to duplicate a sheet including background picture in Aspose.Cells .NET | retain original page layout when adding a worksheet copy using Aspose.Cells | Aspose.Cells AddCopy method background image stream example
// Tags: worksheet AddCopy preserving layout Aspose.Cells | page setup retention during sheet copy .NET | clone sheet visual styling C# | duplicate Excel worksheet layout Aspose.Cells | background picture stream handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads a source workbook, uses the AddCopy method to clone the first worksheet (including its background image), retains the original page‑setup configuration, and saves the result to a new file while handling possible exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "Source.xlsx";
            const string resultPath = "Result.xlsx";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(sourcePath);

            // Clone the first worksheet (including its background image)
            Worksheet sourceSheet = workbook.Worksheets[0];
            int clonedIndex = workbook.Worksheets.AddCopy(sourceSheet.Name);
            Worksheet clonedSheet = workbook.Worksheets[clonedIndex];

            // If explicit background image handling is needed, it can be added here
            // using clonedSheet.SetBackgroundImage(stream) etc.

            // Save the workbook with the cloned worksheet
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved successfully to '{resultPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
