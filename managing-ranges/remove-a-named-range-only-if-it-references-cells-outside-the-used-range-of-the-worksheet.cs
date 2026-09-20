// Title: Remove named ranges that reference cells outside the used range of each worksheet using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that iterates all worksheets, identifies defined names whose range exceeds the worksheet's MaxDataRow/MaxDataColumn, and deletes those names from the NameCollection. | Show a complete example that collects out‑of‑range named ranges, removes them safely, and saves the cleaned workbook.
// Common Searches: Aspose.Cells C# delete defined names that point beyond the used area of a worksheet | how to clean up Excel named ranges that are outside the data region using .NET | remove worksheet‑scoped named ranges that reference empty rows or columns in Aspose.Cells | C# code to filter invalid named ranges in an Excel file with Aspose.Cells
// Tags: remove out-of-range named ranges Aspose.Cells | filter defined names by used range .NET | delete invalid Excel named ranges C# | iterate worksheets to clean NameCollection Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook, determines the used range for each worksheet, finds any defined names whose referenced cells lie outside that range, removes those names from the workbook's NameCollection, and saves the updated file.
class RemoveOutOfRangeNamedRanges
{
    static void Main()
    {
        // Paths for input and output workbooks
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

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

            // Access the collection of defined names (both workbook‑ and worksheet‑scoped)
            NameCollection allNames = workbook.Worksheets.Names;

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine the used range limits of the current worksheet
                int maxDataRow = sheet.Cells.MaxDataRow;          // zero‑based index of last used row
                int maxDataColumn = sheet.Cells.MaxDataColumn;    // zero‑based index of last used column

                // Collect names that need to be removed (cannot modify collection while iterating)
                List<string> namesToRemove = new List<string>();

                // Examine all names
                foreach (Name name in allNames)
                {
                    try
                    {
                        // Get the range the name refers to
                        Aspose.Cells.Range range = name.GetRange();

                        // If the name does not refer to a range, skip it
                        if (range == null)
                            continue;

                        // Ensure the range belongs to the current worksheet
                        if (range.Worksheet != sheet)
                            continue;

                        // Calculate the boundaries of the named range
                        int firstRow = range.FirstRow;
                        int firstCol = range.FirstColumn;
                        int lastRow = firstRow + range.RowCount - 1;
                        int lastCol = firstCol + range.ColumnCount - 1;

                        // Determine if any part of the range lies outside the used range
                        bool outsideUsedRange = firstRow > maxDataRow ||
                                                firstCol > maxDataColumn ||
                                                lastRow > maxDataRow ||
                                                lastCol > maxDataColumn;

                        // Mark the name for removal if it references cells outside the used range
                        if (outsideUsedRange)
                        {
                            // Use the Text property to get the defined name string
                            namesToRemove.Add(name.Text);
                        }
                    }
                    catch (Exception ex)
                    {
                        // Log and continue with other names
                        Console.WriteLine($"Error processing name '{name.Text}': {ex.Message}");
                    }
                }

                // Remove the identified names from the collection
                foreach (string nameToRemove in namesToRemove)
                {
                    allNames.Remove(nameToRemove);
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
