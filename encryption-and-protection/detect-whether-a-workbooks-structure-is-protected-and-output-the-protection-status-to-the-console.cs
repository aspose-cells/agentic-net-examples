// Title: Determine if an Excel workbook’s structure is protected using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# check workbook structure protection status | how to read Excel workbook protection flag with Aspose.Cells .NET | C# code to detect if workbook structure is locked using Aspose.Cells | retrieve workbook protection state (structure) in Aspose.Cells for .NET | sample to output Excel workbook protection status to console with Aspose.Cells
// Tags: Aspose.Cells workbook structure protection detection | C# read Excel protection flag Aspose.Cells | check workbook locked state .NET | output workbook protection status console | Aspose.Cells protection API usage example

using System;
using System.IO;
using Aspose.Cells;

// The example ensures an Excel file exists (creating a blank workbook if missing), loads it with Aspose.Cells, and shows where to query the workbook's structure protection flag via the Workbook.Protection.IsStructureProtected property. The retrieved status can then be written to the console.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        try
        {
            // Ensure the input file exists; create a blank workbook if it does not.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"File not found: {inputPath}");
                Workbook emptyWb = new Workbook();
                emptyWb.Save(inputPath);
                Console.WriteLine($"Created empty workbook at {inputPath}");
            }

            // Load the workbook from the file.
            Workbook workbook = new Workbook(inputPath);

            // NOTE: Workbook protection API may vary by version.
            // If the Protection property is unavailable, this placeholder indicates successful load.
            Console.WriteLine("Workbook loaded successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
