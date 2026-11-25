// Title: Log the time required to enumerate all initialized cells in every worksheet of a large Excel workbook using Aspose.Cells for .NET
// Common Searches: how to benchmark cell enumeration speed with Aspose.Cells in C# | measure time to count initialized cells across worksheets using Aspose.Cells .NET | C# code to log elapsed time while iterating through a large Excel workbook with Aspose.Cells | performance testing of Aspose.Cells Cells collection enumeration in a big workbook | using Stopwatch to time cell counting in Aspose.Cells workbook processing
// Tags: Aspose.Cells cell enumeration performance | initialized cells counting Aspose.Cells | worksheet iteration timing Aspose.Cells | large Excel workbook processing Aspose.Cells | C# Stopwatch Aspose.Cells performance

using System;
using System.Diagnostics;
using Aspose.Cells;

// The example loads a specified Excel file with Aspose.Cells, iterates through each worksheet's Cells collection (which returns only initialized cells), counts the cells, and uses a Stopwatch to capture the total enumeration time. It then outputs the workbook path, worksheet count, total initialized cells, and the elapsed time in seconds and milliseconds.
class Program
{
    static void Main(string[] args)
    {
        // Determine the workbook file path.
        string workbookPath;
        if (args.Length > 0 && !string.IsNullOrWhiteSpace(args[0]))
        {
            workbookPath = args[0];
        }
        else
        {
            // Default path – adjust as needed for your environment.
            workbookPath = "largeWorkbook.xlsx";
        }

        if (!System.IO.File.Exists(workbookPath))
        {
            Console.Error.WriteLine($"Error: Workbook file not found at '{workbookPath}'.");
            return;
        }

        try
        {
            // Load the workbook.
            using (Workbook workbook = new Workbook(workbookPath))
            {
                long totalInitializedCells = 0;
                Stopwatch sw = Stopwatch.StartNew();

                // Iterate through each worksheet.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // The Cells collection implements IEnumerable<Cell> and enumerates only initialized cells.
                    foreach (Cell cell in sheet.Cells)
                    {
                        totalInitializedCells++;
                        // No further processing required; just counting.
                    }
                }

                sw.Stop();

                Console.WriteLine($"Workbook: {workbookPath}");
                Console.WriteLine($"Worksheets processed: {workbook.Worksheets.Count}");
                Console.WriteLine($"Total initialized cells enumerated: {totalInitializedCells}");
                Console.WriteLine($"Time taken: {sw.Elapsed.TotalSeconds:F3} seconds ({sw.ElapsedMilliseconds} ms)");
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An exception occurred while processing the workbook: {ex.Message}");
            Console.Error.WriteLine(ex.StackTrace);
        }
    }
}
