// Title: Clear the AutoFilter on column AJ and display all rows using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# remove autofilter from column AJ and show all rows | How to clear a specific column filter in an Excel file with Aspose.Cells .NET | Refresh worksheet after deleting AutoFilter criteria using Aspose.Cells C# | Programmatically delete Excel AutoFilter on column index 35 with Aspose.Cells | C# code to unfilter Excel sheet and make hidden rows visible using Aspose.Cells
// Tags: autofilter remove column aj aspnet | worksheet refresh after filter clear aspnet | aspnet cells delete column filter by index | excel unfilter show all rows c# | clear specific autofilter criteria aspnet

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, accesses the first worksheet, obtains its AutoFilter object, removes any filter applied to column AJ (zero‑based index 35), refreshes the filter to make all rows visible again, and saves the modified workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths
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
                // Load the workbook from the input file
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (index 0)
                Worksheet worksheet = workbook.Worksheets[0];

                // Get the AutoFilter object for the worksheet
                AutoFilter autoFilter = worksheet.AutoFilter;

                // If an AutoFilter exists, remove filter for column AJ (zero‑based index 35)
                if (autoFilter != null)
                {
                    // Remove any existing filter on the specified column
                    autoFilter.RemoveFilter(35);
                    // Refresh to make all rows visible again
                    autoFilter.Refresh();
                }

                // Save the modified workbook to the output file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
