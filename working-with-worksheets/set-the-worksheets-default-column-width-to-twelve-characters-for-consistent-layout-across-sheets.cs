// Title: Set a default column width of 12 characters for every column in an Aspose.Cells worksheet using C#
// AI Prompts: Generate C# code that uses Aspose.Cells to set the width of all columns in the first worksheet to 12 characters and then saves the file. | Provide a C# snippet that iterates over the first 256 columns of a worksheet and applies a uniform column width of 12 characters with Aspose.Cells.
// Common Searches: Aspose.Cells C# set column width for entire worksheet | How to define default column width of 12 characters in Aspose.Cells workbook | C# loop to apply same column width to multiple columns using Aspose.Cells | Set uniform column width across all columns in an Aspose.Cells sheet | Default column width setting in Aspose.Cells .NET API
// Tags: set column width Aspose.Cells C# | default column width 12 characters Aspose.Cells | iterate columns set width worksheet Aspose.Cells | uniform column width workbook .NET | save workbook after column width adjustment Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The program creates a new workbook, accesses the first worksheet, loops through the first 256 columns to set each column's width to 12 characters, and saves the workbook as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set column width for a reasonable range (e.g., first 256 columns) to emulate a default width
            const double columnWidth = 12;
            for (int col = 0; col < 256; col++)
            {
                sheet.Cells.SetColumnWidth(col, columnWidth);
            }

            // Define output file path
            string outputPath = "Output.xlsx";

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
