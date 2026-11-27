// Title: Setting worksheet TabId to zero and handling the resulting exception with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set worksheet TabId to 0 and catch exception | How to handle exception when assigning zero TabId to worksheets in Aspose.Cells | Iterate all worksheets and log expected TabId zero error using Aspose.Cells .NET | C# example for catching Aspose.Cells TabId property exception | Why does setting TabId = 0 throw in Aspose.Cells and how to log it
// Tags: Aspose.Cells worksheet TabId exception handling | C# set TabId zero Aspose.Cells | iterate worksheets catch TabId error Aspose.Cells | log expected exception Aspose.Cells TabId | save workbook after TabId operation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTabIdDemo
{
    // The sample loads an existing workbook or creates a new one, iterates through each worksheet, attempts to assign TabId = 0 (which triggers an exception), catches and logs the expected exception for every sheet, and finally saves the workbook to confirm that other operations succeed.
    class Program
    {
        static void Main(string[] args)
        {
            // Input workbook path (optional). If the file does not exist, a new workbook will be created.
            string inputPath = "input.xlsx";

            Workbook workbook;

            if (File.Exists(inputPath))
            {
                try
                {
                    workbook = new Workbook(inputPath);
                    Console.WriteLine($"Loaded workbook from '{inputPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to load workbook from '{inputPath}'. Exception: {ex.GetType().Name} - {ex.Message}");
                    // Fallback to a new workbook if loading fails
                    workbook = new Workbook();
                    Console.WriteLine("Created a new workbook as fallback.");
                }
            }
            else
            {
                // Create a new workbook with a few worksheets for demonstration
                workbook = new Workbook();
                workbook.Worksheets.Add("Sheet2");
                workbook.Worksheets.Add("Sheet3");
                Console.WriteLine("Input file not found. Created a new workbook with default sheets.");
            }

            // Iterate through all worksheets and attempt to set TabId to zero
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                try
                {
                    // Attempt to set TabId to zero – this is expected to throw an exception
                    worksheet.TabId = 0;

                    // If no exception is thrown (unlikely), log the successful assignment
                    Console.WriteLine($"[SUCCESS] Worksheet '{worksheet.Name}': TabId set to {worksheet.TabId}.");
                }
                catch (Exception ex)
                {
                    // Log the expected exception details
                    Console.WriteLine($"[EXPECTED EXCEPTION] Worksheet '{worksheet.Name}': Unable to set TabId to zero. Exception: {ex.GetType().Name} - {ex.Message}");
                }
            }

            // Save the workbook to verify that other operations succeeded
            string outputPath = "output.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook. Exception: {ex.GetType().Name} - {ex.Message}");
            }

            // Keep console window open when run outside of an IDE
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
    }
}
