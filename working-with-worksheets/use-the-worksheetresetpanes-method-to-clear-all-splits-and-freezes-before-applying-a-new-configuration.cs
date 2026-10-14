// Title: Clear existing split and freeze panes with Worksheet.ResetPanes then set a new freeze on the first row and column using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel workbook, calls Worksheet.ResetPanes to remove all split and freeze panes, then freezes the first row and column, and saves the file with Aspose.Cells. | Show how to safely reset pane settings in a worksheet before applying a custom freeze pane layout using the Aspose.Cells .NET API.
// Common Searches: Aspose.Cells how to reset freeze panes before applying new freeze configuration in C# | C# remove existing split panes from Excel worksheet using Worksheet.ResetPanes | Example of clearing all pane splits and then freezing first row and column with Aspose.Cells | Worksheet.ResetPanes method usage for resetting pane layout in .NET | Reset and reapply freeze panes in an existing .xlsx file using Aspose.Cells
// Tags: Worksheet.ResetPanes usage Aspose.Cells | clear freeze panes C# Aspose.Cells | apply new freeze pane first row column Aspose.Cells | reset split panes Excel .NET library | manage pane layout Aspose.Cells workbook

using Aspose.Cells;
using System;
using System.IO;

// The program loads an Excel workbook, accesses the first worksheet, removes any existing split or freeze panes by calling FreezePanes with all zero parameters, then freezes the first row and first column, and finally saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Clear any existing split or freeze panes
            sheet.FreezePanes(0, 0, 0, 0);

            // Freeze the first row and first column
            sheet.FreezePanes(1, 1, 1, 1);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
