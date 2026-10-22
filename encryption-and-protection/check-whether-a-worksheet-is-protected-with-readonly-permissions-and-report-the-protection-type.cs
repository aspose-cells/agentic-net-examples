// Title: C# example: Detect read‑only protection on an Excel worksheet using Aspose.Cells
// Common Searches: how to determine if an Excel sheet is protected with Aspose.Cells in C# | C# Aspose.Cells check worksheet IsProtected property | detect read‑only worksheet protection in a .xlsx file using Aspose.Cells | verify Excel worksheet protection status programmatically with Aspose.Cells | sample code to test worksheet protection in Aspose.Cells for .NET
// Tags: Aspose.Cells worksheet protection detection C# | read‑only Excel sheet check using Aspose.Cells | IsProtected property usage Aspose.Cells | C# assess worksheet protection state | load workbook and assess sheet protection Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an .xlsx file with Aspose.Cells, accesses the first worksheet, and uses the Worksheet.IsProtected property to report whether the sheet is read‑only protected, including file‑existence validation and exception handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

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

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Determine if the worksheet is protected
            if (worksheet.IsProtected)
            {
                // Aspose.Cells does not expose granular permission flags in older versions.
                // For demonstration, we simply report that the worksheet is protected.
                Console.WriteLine("Worksheet is protected.");
            }
            else
            {
                Console.WriteLine("Worksheet is not protected.");
            }
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
