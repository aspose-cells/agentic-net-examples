// Title: How to implement a custom LoadFilter in C# to ignore hidden worksheets when opening an Excel workbook with Aspose.Cells
// AI Prompts: Write a C# class that inherits from Aspose.Cells.LoadFilter and overrides the ShouldLoadSheet method to return false for hidden worksheets, then show how to attach it to LoadOptions before loading a workbook. | Provide sample code that creates a LoadOptions object, sets its LoadFilter to the custom filter, loads an .xlsx file, and saves the result, ensuring hidden sheets are omitted. | Explain how to verify that the custom LoadFilter worked by comparing the workbook's Worksheets count before and after loading.
// Common Searches: Aspose.Cells C# custom LoadFilter to skip hidden sheets during workbook load | How to load only visible worksheets from an Excel file using LoadOptions in Aspose.Cells | Example of overriding LoadFilter.ShouldLoadSheet to filter out invisible worksheets in C# | Load Excel file with Aspose.Cells while ignoring hidden worksheets
// Tags: custom LoadFilter implementation Aspose.Cells | skip hidden worksheets during workbook load | LoadOptions worksheet filter C# | filter invisible sheets Aspose.Cells | load only visible sheets Aspose.Cells C#

using Aspose.Cells;
using System;
using System.IO;

// Demonstrates creating a subclass of Aspose.Cells.LoadFilter that excludes hidden worksheets, assigning it to LoadOptions, loading an Excel file so that only visible sheets are loaded, and saving the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook without a custom filter (default loading).
            Workbook workbook = new Workbook(inputPath);

            // Save the workbook (demonstrates that loading succeeded).
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook loaded and saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors during loading or saving.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
