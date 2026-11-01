// Title: How to verify that moving cells H1:H5 to I1:I5 won't overwrite existing data using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# check if destination range I1:I5 is empty before moving cells | C# validate no data overlap when copying range H1:H5 to I1:I5 with Aspose.Cells | detect existing values in target cells before moving a range in Aspose.Cells .NET | prevent cell overwrite when shifting columns using Aspose.Cells in C# | ensure safe range move without collision in an Aspose.Cells workbook
// Tags: Aspose.Cells range overlap validation | C# check empty destination cells Aspose.Cells | prevent data collision moving cells Aspose.Cells | validate range move safety .NET | Aspose.Cells detect existing cell values

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, creates source (H1:H5) and destination (I1:I5) ranges, iterates through the destination cells to detect any non‑null values, flags an overlap if data exists, and reports whether the target range is empty and safe for moving.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Use fully qualified type to avoid ambiguity with System.Range
            Aspose.Cells.Range sourceRange = worksheet.Cells.CreateRange("H1", "H5");
            Aspose.Cells.Range destinationRange = worksheet.Cells.CreateRange("I1", "I5");

            // Validate that the destination range does not contain existing data
            bool overlapDetected = false;
            foreach (Cell cell in destinationRange)
            {
                // If any cell in the destination range has a non‑null value, overlap exists
                if (cell.Value != null)
                {
                    overlapDetected = true;
                    break;
                }
            }

            // Output the validation result
            if (overlapDetected)
            {
                Console.WriteLine("Overlap detected: destination range I1:I5 contains data.");
            }
            else
            {
                Console.WriteLine("No overlap: destination range I1:I5 is empty and safe for moving.");
            }

            // Optional: save the workbook if modifications are made
            // workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
