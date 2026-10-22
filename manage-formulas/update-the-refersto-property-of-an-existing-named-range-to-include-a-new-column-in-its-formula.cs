// Title: How to expand an existing named range's RefersTo formula to include an additional column using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells add column D to an existing named range RefersTo | update named range formula to extend range in Excel with Aspose.Cells .NET | programmatically modify RefersTo property of a named range in a workbook using Aspose.Cells | extend Excel named range to new column while preserving row count using Aspose.Cells C# | parse and rebuild RefersTo string for a named range with Aspose.Cells example
// Tags: Aspose.Cells named range RefersTo modification | C# extend Excel named range column | Aspose.Cells workbook range expansion | update RefersTo property programmatically | Excel named range column addition .NET

using Aspose.Cells;
using System;
using System.IO;
using System.Text.RegularExpressions;

namespace MyApp
{
    // The example loads an Excel workbook, retrieves the named range "MyRange", parses its RefersTo string to identify the sheet, start cell, and end row, constructs a new end cell in column D while keeping the same row number, rebuilds the RefersTo expression to include the new column, assigns the updated RefersTo back to the named range, and saves the workbook.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            try
            {
                // Load the workbook (no password supplied)
                Workbook workbook = new Workbook(inputPath);

                // Retrieve the named range by its name
                Name namedRange = workbook.Worksheets.Names["MyRange"];
                if (namedRange == null)
                {
                    Console.WriteLine("Named range 'MyRange' not found.");
                    return;
                }

                // Example RefersTo: Sheet1!$A$1:$C$10
                string oldRefersTo = namedRange.RefersTo;

                // Split the RefersTo into sheet part and range part
                int exclPos = oldRefersTo.IndexOf('!');
                if (exclPos < 0)
                {
                    Console.WriteLine("Invalid RefersTo format.");
                    return;
                }

                string sheetPart = oldRefersTo.Substring(0, exclPos);          // e.g., Sheet1
                string rangePart = oldRefersTo.Substring(exclPos + 1);         // e.g., $A$1:$C$10

                // Separate start and end cells of the range
                string[] cells = rangePart.Split(':');
                if (cells.Length != 2)
                {
                    Console.WriteLine("RefersTo does not contain a start and end cell.");
                    return;
                }

                string startCell = cells[0];   // e.g., $A$1
                string endCell = cells[1];     // e.g., $C$10

                // Extract the row number from the end cell
                string rowNumber = Regex.Match(endCell, @"\d+").Value; // e.g., 10
                if (string.IsNullOrEmpty(rowNumber))
                {
                    Console.WriteLine("Unable to extract row number from end cell.");
                    return;
                }

                // Define the new column to add (e.g., column D)
                string newEndCell = "$D$" + rowNumber; // e.g., $D$10

                // Construct the new RefersTo string
                string newRefersTo = $"{sheetPart}!{startCell}:{newEndCell}";

                // Update the RefersTo property of the named range
                namedRange.RefersTo = newRefersTo;

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (CellsException ex)
            {
                // Handle password-protected workbook or other Cells-specific errors
                Console.WriteLine($"Aspose.Cells error: {ex.Message}");
            }
            catch (Exception ex)
            {
                // General exception handling
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
