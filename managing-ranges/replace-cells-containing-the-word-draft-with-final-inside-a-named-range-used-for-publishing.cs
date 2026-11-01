// Title: Replace cells containing "Draft" with "Final" in a named range using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# replace specific text in a named range of an Excel file | How to change cells that contain 'Draft' to 'Final' within a defined named range using Aspose.Cells | Iterating over cells in a named range to update values with Aspose.Cells for .NET | Saving an Excel workbook after editing values in a named range with C# Aspose.Cells | C# code example for updating string values in a named range using Aspose.Cells library
// Tags: named range cell iteration Aspose.Cells | replace cell string value .NET Excel | update Excel named range values C# | modify workbook cells based on content Aspose.Cells | save workbook after named range edit .NET

using System;
using System.IO;
using Aspose.Cells;

namespace ReplaceDraftWithFinalApp
{
    // The example loads an existing workbook, retrieves the named range "PublishRange", iterates through each cell in that range, replaces any cell whose text equals "Draft" with "Final", and then saves the modified workbook to a new file.
    class ReplaceDraftWithFinal
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Ensure the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Retrieve the named range (must exist in the workbook)
                Aspose.Cells.Range publishRange = workbook.Worksheets.GetRangeByName("PublishRange");
                if (publishRange == null)
                {
                    Console.WriteLine("Named range 'PublishRange' not found.");
                    return;
                }

                // Replace "Draft" with "Final" within the named range
                foreach (Cell cell in publishRange)
                {
                    if (cell.StringValue == "Draft")
                    {
                        cell.PutValue("Final");
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
