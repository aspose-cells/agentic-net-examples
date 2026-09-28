// Title: Load an existing Excel workbook, update a named range, and save as a new file using Aspose.Cells for .NET (C#)
// AI Prompts: Using Aspose.Cells in C#, open 'Report.xlsx', locate the named range 'MyRange', set the first cell's value to 'Updated', and write the workbook to 'ReportUpdated.xlsx'. | Retrieve a named range by name from a workbook with Aspose.Cells, iterate over its cells to assign a new value, and save the modified workbook under a different filename in C#. | Add error handling to verify that 'Report.xlsx' exists and that the named range is present before updating its cells and saving the updated file with Aspose.Cells.
// Common Searches: asp.net c# how to change a specific named range in an existing Excel file with Aspose.Cells | example code to load workbook, edit named range, and save as new file using Aspose.Cells for .NET | retrieve named range by name and update cell values in C# Aspose.Cells | save modified Excel workbook after editing named range with Aspose.Cells library
// Tags: load workbook with Aspose.Cells C# | retrieve named range by name Aspose.Cells | modify cell value in named range Aspose.Cells | save updated workbook Aspose.Cells C# | error handling missing file Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // // C# program that uses Aspose.Cells to open 'Report.xlsx', fetches the named range 'MyRange', updates the first cell to "Updated" (optionally iterates all cells), and saves the result as 'ReportUpdated.xlsx' with basic file‑existence and range‑presence checks.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputFile = "Report.xlsx";
                const string outputFile = "ReportUpdated.xlsx";
                const string namedRangeName = "MyRange";

                // Verify that the input workbook exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file \"{inputFile}\" not found.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputFile);

                // Retrieve the range associated with the named range
                Aspose.Cells.Range range = workbook.Worksheets.GetRangeByName(namedRangeName);

                if (range != null)
                {
                    // Example modification: set the value of the first cell in the range
                    range[0, 0].PutValue("Updated");

                    // Uncomment the following block to modify all cells in the range
                    // for (int row = 0; row < range.RowCount; row++)
                    // {
                    //     for (int col = 0; col < range.ColumnCount; col++)
                    //     {
                    //         range[row, col].PutValue("Updated");
                    //     }
                    // }
                }
                else
                {
                    Console.WriteLine($"Named range \"{namedRangeName}\" not found.");
                }

                // Save the modified workbook
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved as \"{outputFile}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
