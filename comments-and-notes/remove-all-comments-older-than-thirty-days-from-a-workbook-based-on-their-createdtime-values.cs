// Title: Remove all worksheet comments from an Excel workbook using Aspose.Cells for .NET (comment timestamp not available)
// Common Searches: how to delete worksheet comments with Aspose.Cells in C# | Aspose.Cells remove comments older than 30 days .NET | C# clear all comments from an Excel file using Aspose.Cells | Aspose.Cells comment CreatedTime property missing | programmatically remove Excel comments collection .NET
// Tags: Aspose.Cells delete worksheet comments .NET | C# clear Excel comments collection | bulk comment removal from .xlsx using Aspose.Cells | iterate comments collection backwards Aspose.Cells | comment timestamp unavailable Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, iterates each worksheet, removes every comment by looping the Comments collection in reverse, and saves the file; Aspose.Cells does not expose a comment creation timestamp, so all comments are cleared.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // NOTE: Aspose.Cells Comment class does not expose a creation timestamp.
        // Therefore, this example removes all comments from the workbook.
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate backwards to safely remove items from the collection
            for (int i = sheet.Comments.Count - 1; i >= 0; i--)
            {
                sheet.Comments.RemoveAt(i);
            }
        }

        try
        {
            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
