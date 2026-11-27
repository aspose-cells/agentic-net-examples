// Title: Remove all horizontal and vertical page breaks from an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells delete all page breaks from an Excel workbook | how to clear horizontal and vertical page breaks using Aspose.Cells .NET | reset pagination in Excel by removing page breaks with Aspose.Cells | programmatically clear page break collections for each worksheet in C# | Aspose.Cells remove page breaks before saving workbook
// Tags: page break collection removal Aspose.Cells | horizontal pagination break removal Aspose.Cells | vertical pagination break removal Aspose.Cells | worksheet pagination control Aspose.Cells | save workbook after pagination reset Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace RemovePageBreaksDemo
{
    // The C# program loads an Excel workbook with Aspose.Cells, iterates through every worksheet, clears both horizontal and vertical page break collections, and saves the updated workbook, enabling automatic pagination.
    class Program
    {
        static void Main(string[] args)
        {
            // Validate input arguments
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: RemovePageBreaksDemo <input.xlsx> <output.xlsx>");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];

            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" does not exist.");
                return;
            }

            try
            {
                // Load the workbook from the specified file
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Remove all horizontal page breaks
                    HorizontalPageBreakCollection hBreaks = sheet.HorizontalPageBreaks;
                    if (hBreaks != null && hBreaks.Count > 0)
                    {
                        hBreaks.Clear();
                        Console.WriteLine($"Removed {hBreaks.Count} horizontal page breaks from sheet \"{sheet.Name}\".");
                    }

                    // Remove all vertical page breaks
                    VerticalPageBreakCollection vBreaks = sheet.VerticalPageBreaks;
                    if (vBreaks != null && vBreaks.Count > 0)
                    {
                        vBreaks.Clear();
                        Console.WriteLine($"Removed {vBreaks.Count} vertical page breaks from sheet \"{sheet.Name}\".");
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
