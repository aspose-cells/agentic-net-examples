// Title: Automatically adjust Excel column widths to fit cell contents using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells auto fit column width based on cell values | How to auto‑size all columns in an existing workbook with Aspose.Cells .NET | Aspose.Cells programmatically set column width to content length in C# | Auto‑fit Excel columns after loading file using Aspose.Cells for .NET | Adjust worksheet column widths to fit data with Aspose.Cells C# example
// Tags: auto-fit column Aspose.Cells | auto-size Excel columns C# | worksheet column width adjustment Aspose.Cells | auto-fit all columns .NET | adjust column width based on content C#

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an existing Excel workbook, checks the file's existence, iterates through each used column in the first worksheet, calls AutoFitColumn to automatically size columns to their content, and saves the result to a new file while handling potential password‑protected file errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook (no password supplied)
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the last used column
            int lastUsedColumn = sheet.Cells.MaxColumn;

            // Auto‑fit each column that contains data
            for (int col = 0; col <= lastUsedColumn; col++)
            {
                sheet.AutoFitColumn(col);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle password‑protected files or other errors
            if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine("The workbook is password protected. Please provide a valid password.");
            }
            else
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
