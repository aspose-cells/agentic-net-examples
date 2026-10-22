// Title: Load an Excel workbook from a UNC network share using Aspose.Cells Workbook constructor in C#
// Common Searches: C# Aspose.Cells load Excel file from UNC path with file path constructor | How to open a workbook located on a network share using Aspose.Cells in .NET | Check file existence before creating Aspose.Cells Workbook from UNC location | Reading first worksheet name from Excel file on a network share with Aspose.Cells | Handling UNC path errors when loading Excel workbook with Aspose.Cells C#
// Tags: Aspose.Cells workbook constructor UNC path | C# load Excel from network share | verify file existence Aspose.Cells | read first worksheet name Aspose.Cells | handle UNC path errors Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates how to verify a UNC network‑share path, load the Excel file with the Aspose.Cells Workbook constructor, retrieve the name of the first worksheet, and gracefully handle any exceptions that may occur.
class Program
{
    static void Main()
    {
        // UNC path to the workbook on a network share
        string uncPath = @"\\ServerName\ShareFolder\MyWorkbook.xlsx";

        // Verify that the file exists before attempting to load
        if (!File.Exists(uncPath))
        {
            Console.WriteLine($"File not found: {uncPath}");
            return;
        }

        try
        {
            // Load the workbook using the file path constructor (handles UNC correctly)
            Workbook workbook = new Workbook(uncPath);

            // Output the name of the first worksheet, if any
            if (workbook.Worksheets.Count > 0)
            {
                Console.WriteLine("First worksheet: " + workbook.Worksheets[0].Name);
            }
            else
            {
                Console.WriteLine("The workbook contains no worksheets.");
            }
        }
        catch (Exception ex)
        {
            // Handle any errors that occur during loading or processing
            Console.WriteLine("An error occurred while processing the workbook:");
            Console.WriteLine(ex.Message);
        }
    }
}
