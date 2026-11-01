// Title: Calculate the total number of threaded comments in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells get total threaded comments in an .xlsx file | How to count all threaded comments across worksheets with Aspose.Cells .NET | Sum of ThreadedComments collection in each sheet using Aspose.Cells | Retrieve number of threaded comments in Excel workbook programmatically in C# | Aspose.Cells count threaded comments per worksheet and aggregate
// Tags: Aspose.Cells ThreadedComments collection enumeration | C# aggregate workbook threaded comment count | reflection access to ThreadedComments property in Aspose.Cells | Excel .xlsx threaded comment total calculation .NET | iterate worksheets for comment aggregation using Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel file with Aspose.Cells, iterates through every worksheet, uses reflection to obtain the ThreadedComments collection when available, adds up each collection's Count, and prints the workbook‑wide total of threaded comments while handling missing files and per‑sheet errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Ensure the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            int totalThreadedComments = 0;

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    // Use reflection to access ThreadedComments (may not exist in older versions)
                    var threadedCommentsProp = sheet.GetType().GetProperty("ThreadedComments");
                    if (threadedCommentsProp != null)
                    {
                        var threadedComments = threadedCommentsProp.GetValue(sheet);
                        if (threadedComments != null)
                        {
                            var countProp = threadedComments.GetType().GetProperty("Count");
                            if (countProp != null)
                            {
                                int count = (int)countProp.GetValue(threadedComments);
                                totalThreadedComments += count;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Log worksheet‑specific errors but continue processing other sheets
                    Console.WriteLine($"Warning: Could not retrieve threaded comments for sheet \"{sheet.Name}\": {ex.Message}");
                }
            }

            // Output the total number of threaded comments in the workbook
            Console.WriteLine("Total threaded comments: " + totalThreadedComments);
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors during workbook processing
            Console.WriteLine("An error occurred while processing the workbook: " + ex.Message);
        }
    }
}
