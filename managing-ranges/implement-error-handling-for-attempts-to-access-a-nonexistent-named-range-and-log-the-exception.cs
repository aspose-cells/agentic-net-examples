using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class NamedRangeErrorHandlingDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add a valid named range for demonstration
                int existingIndex = workbook.Worksheets.Names.Add("ExistingRange");
                workbook.Worksheets.Names[existingIndex].RefersTo = "=Sheet1!$A$1:$A$3";

                // Attempt to access a non‑existent named range and handle the error
                try
                {
                    // Retrieve the name object; returns null if the name does not exist
                    Name missingName = workbook.Worksheets.Names["MissingRange"];

                    // Throw if not found
                    if (missingName == null)
                        throw new InvalidOperationException("Named range 'MissingRange' does not exist.");

                    // Get its range (won't be reached in this demo)
                    Aspose.Cells.Range missingRange = missingName.GetRange();
                    Console.WriteLine("Missing range address: " + missingRange.Address);
                }
                catch (Exception ex)
                {
                    // Log the exception details
                    Console.WriteLine($"Error accessing named range: {ex.Message}");
                }

                // Save the workbook
                string outputPath = "NamedRangeErrorHandlingDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            NamedRangeErrorHandlingDemo.Run();
        }
    }
}