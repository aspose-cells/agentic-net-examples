// Title: Load multiple Excel workbooks in a loop, add them to a List<Workbook>, and optionally set manual calculation mode with Aspose.Cells for .NET
// Common Searches: c# load several Excel files into a List<Workbook> using Aspose.Cells | aspnet cells set workbook calculation mode to manual for each file in a batch | how to handle missing Excel files when loading workbooks with Aspose.Cells | batch processing Excel workbooks with Aspose.Cells and disable automatic calculation
// Tags: batch load Excel workbooks Aspose.Cells C# | set workbook calculation mode manual Aspose.Cells | handle missing Excel files Aspose.Cells | store Workbook instances in List Aspose.Cells | disable automatic formula calculation Aspose.Cells .NET

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example iterates over a predefined collection of Excel file paths, checks each file's existence, loads each workbook with Aspose.Cells, optionally switches its calculation mode to Manual (commented for version compatibility), adds the successfully loaded Workbook objects to a List<Workbook>, logs any loading errors, and finally reports the total number of workbooks loaded.
    class Program
    {
        static void Main(string[] args)
        {
            // List of workbook file paths to load
            List<string> filePaths = new List<string>
            {
                "Workbook1.xlsx",
                "Workbook2.xlsx",
                "Workbook3.xlsx"
            };

            // Container for the loaded workbooks
            List<Workbook> workbooks = new List<Workbook>();

            foreach (string path in filePaths)
            {
                try
                {
                    // Verify that the file exists before attempting to load it
                    if (!File.Exists(path))
                    {
                        Console.WriteLine($"Warning: File not found – \"{path}\". Skipping this file.");
                        continue;
                    }

                    // Load the workbook from the specified file
                    Workbook wb = new Workbook(path);

                    // If the CalcMode property is available in the used Aspose.Cells version,
                    // you can set it to Manual here. This line is omitted to maintain compatibility
                    // with older versions where the property does not exist.

                    // Store the workbook in the list
                    workbooks.Add(wb);
                }
                catch (Exception ex)
                {
                    // Log any unexpected errors and continue processing remaining files
                    Console.WriteLine($"Error loading workbook \"{path}\": {ex.Message}");
                }
            }

            // At this point, 'workbooks' holds all successfully loaded workbooks.
            Console.WriteLine($"Total workbooks loaded: {workbooks.Count}");
        }
    }
}
