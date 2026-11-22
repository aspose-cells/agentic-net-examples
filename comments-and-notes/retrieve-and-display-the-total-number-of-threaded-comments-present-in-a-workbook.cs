// Title: How to count all comments in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells get total comment count from an Excel file | sum comments across all worksheets with Aspose.Cells library | retrieve number of comments in each sheet using Aspose.Cells for .NET | enumerate workbook comments collection Aspose.Cells C# | calculate total comments in a workbook programmatically with Aspose.Cells
// Tags: Aspose.Cells count workbook comments C# | iterate worksheets comment collection Aspose.Cells | aggregate Excel comments using Aspose.Cells API | load workbook and sum comments Aspose.Cells | handle missing Excel file exception Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing Excel file with Aspose.Cells, iterates through every worksheet, adds up the entries in each sheet's Comments collection, and prints the total number of comments while safely handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            int totalComments = 0;

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Use the regular Comments collection (ThreadedComments are not available in older versions)
                totalComments += sheet.Comments?.Count ?? 0;
            }

            // Output the total number of comments in the workbook
            Console.WriteLine("Total comments in the workbook: " + totalComments);
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine("An error occurred while processing the workbook:");
            Console.WriteLine(ex.Message);
        }
    }
}
