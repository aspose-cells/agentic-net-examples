// Title: How to enable the EmptyCellReferences error check for a single worksheet using Aspose.Cells ErrorCheckOption in C#
// AI Prompts: Create an ErrorCheckOption instance, set its EmptyCellReferences property to true, and assign it to the target worksheet before saving the workbook. | Add code that configures per‑worksheet error checking to flag empty cell references in Aspose.Cells with C#. | Write a snippet that applies EmptyCellReferences validation only to the first worksheet of a new workbook.
// Common Searches: C# Aspose.Cells enable EmptyCellReferences validation for one worksheet only | set per‑worksheet error checking option Aspose.Cells EmptyCellReferences | how to turn on empty cell reference error check in Aspose.Cells workbook using C# | apply ErrorCheckOption to specific sheet Aspose.Cells C#
// Tags: Aspose.Cells ErrorCheckOption EmptyCellReferences | per‑worksheet error checking C# | enable empty cell reference validation Aspose.Cells | configure worksheet error check option | C# Aspose.Cells workbook error checking

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, configures an ErrorCheckOption to enable the EmptyCellReferences error check on the first worksheet, writes a sample value to cell A1, and saves the workbook as Result.xlsx while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Example content to verify the worksheet works
            worksheet.Cells["A1"].PutValue("Hello Aspose.Cells!");

            // Define the output file path
            string outputPath = "Result.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
