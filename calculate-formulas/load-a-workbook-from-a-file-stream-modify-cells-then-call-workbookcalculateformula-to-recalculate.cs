// Title: Load an Excel workbook from a Stream, modify cells, and recalculate all formulas with Aspose.Cells for .NET
// Common Searches: how to load an Excel file from a FileStream using Aspose.Cells in C# | updating cell values and triggering formula recalculation with Aspose.Cells Workbook.CalculateFormula | saving a modified workbook after changing cells with Aspose.Cells .NET | example code for processing an Excel stream and recalculating formulas programmatically | Aspose.Cells recalculate all formulas after editing cells in C#
// Tags: load workbook from stream Aspose.Cells | modify cell values C# Aspose.Cells | recalculate formulas Workbook.CalculateFormula | save workbook to file Aspose.Cells | process Excel stream .NET

using Aspose.Cells;
using System;
using System.IO;

// This example demonstrates loading an Excel workbook from a Stream using Aspose.Cells for .NET, updating specific cells (e.g., writing text to A1 and a number to B2), invoking Workbook.CalculateFormula to recalculate every formula in the workbook, and then saving the modified file as 'ModifiedWorkbook.xlsx' while handling potential exceptions.
public class WorkbookProcessor
{
    /// <param name="fileStream">Stream containing the Excel file.</param>
    public void ProcessWorkbook(Stream fileStream)
    {
        try
        {
            // Load the workbook from the provided file stream
            Workbook workbook = new Workbook(fileStream);

            // Access the first worksheet (index 0)
            Worksheet sheet = workbook.Worksheets[0];

            // Modify cells – example: write text to A1 and a number to B2
            sheet.Cells["A1"].PutValue("Hello, Aspose!");
            sheet.Cells["B2"].PutValue(12345);

            // Recalculate all formulas in the workbook after the changes
            workbook.CalculateFormula();

            // Save the modified workbook to a new file
            string outputPath = "ModifiedWorkbook.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error processing workbook: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        // Expect input file path as first argument
        if (args.Length == 0)
        {
            Console.Error.WriteLine("Please provide the path to the input Excel file as an argument.");
            return;
        }

        string inputPath = args[0];

        // Verify the input file exists
        if (!File.Exists(inputPath))
        {
            Console.Error.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            using (FileStream fs = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
            {
                WorkbookProcessor processor = new WorkbookProcessor();
                processor.ProcessWorkbook(fs);
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
