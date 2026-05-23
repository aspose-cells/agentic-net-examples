using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace NamedRangeValidator
{
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for input and output workbooks
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            Workbook workbook = null;
            try
            {
                // Load the workbook
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // List to collect names with invalid references
            var invalidNames = new List<string>();

            try
            {
                // Iterate through all defined names in the workbook
                NameCollection names = workbook.Worksheets.Names;
                foreach (Name name in names)
                {
                    try
                    {
                        // Retrieve the ranges the name refers to without recalculation
                        // GetRanges(false) returns null if the reference cannot be resolved
                        AsposeRange[] ranges = name.GetRanges(false);

                        // If no ranges are returned, the reference is invalid
                        if (ranges == null || ranges.Length == 0)
                        {
                            invalidNames.Add($"{name.Text} -> {name.RefersTo}");
                            continue;
                        }

                        // Additional sanity check: ensure each range points to an existing worksheet
                        foreach (AsposeRange rng in ranges)
                        {
                            // The Range object contains a Worksheet reference; if it's null, the sheet does not exist
                            if (rng.Worksheet == null)
                            {
                                invalidNames.Add($"{name.Text} -> {name.RefersTo}");
                                break;
                            }

                            // Verify that the range's start/end rows/columns are within the worksheet's dimensions
                            // (Aspose.Cells automatically expands the worksheet, so this check mainly catches
                            // malformed addresses like $A$0 or negative indices)
                            if (rng.FirstRow < 0 || rng.FirstColumn < 0 ||
                                rng.FirstRow >= rng.Worksheet.Cells.MaxDataRow + 1 ||
                                rng.FirstColumn >= rng.Worksheet.Cells.MaxDataColumn + 1)
                            {
                                invalidNames.Add($"{name.Text} -> {name.RefersTo}");
                                break;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // Any exception while resolving the name is treated as an invalid reference
                        invalidNames.Add($"{name.Text} -> {name.RefersTo} (Error: {ex.Message})");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error while processing named ranges: {ex.Message}");
                return;
            }

            // Report results
            Console.WriteLine("Named Range Validation Report");
            Console.WriteLine("-----------------------------");
            if (invalidNames.Count == 0)
            {
                Console.WriteLine("All named ranges reference valid cells.");
            }
            else
            {
                Console.WriteLine("Invalid named ranges found:");
                foreach (string entry in invalidNames)
                {
                    Console.WriteLine(entry);
                }
            }

            // Save the workbook (unchanged) – ensure the output directory exists
            try
            {
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}