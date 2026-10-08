// Title: How to export an Excel workbook to JSON while skipping hidden rows and columns with Aspose.Cells in C#
// AI Prompts: Write C# code that loads an .xlsx file, configures JsonSaveOptions to ignore hidden rows and columns, and saves the workbook as a JSON file. | Show an example of using Aspose.Cells JsonSaveOptions in C# to export Excel data to JSON without including hidden rows or columns.
// Common Searches: Aspose.Cells C# export Excel to JSON exclude hidden rows and columns | skip hidden rows when saving workbook as JSON using Aspose.Cells | C# JsonSaveOptions hide columns omission example
// Tags: Aspose.Cells JsonSaveOptions hide rows | C# export Excel to JSON without hidden columns | skip hidden cells during JSON conversion Aspose | configure workbook JSON save to omit hidden rows

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing .xlsx workbook, creates a JsonSaveOptions object (which by default omits hidden rows and columns), and saves the workbook as a JSON file using those options.
class Program
{
    static void Main()
    {
        try
        {
            const string inputFile = "input.xlsx";
            const string outputFile = "output.json";

            // Verify that the input workbook exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: The file '{inputFile}' was not found.");
                return;
            }

            // Load the workbook from the existing Excel file
            Workbook workbook = new Workbook(inputFile);

            // Configure JSON save options (default behavior excludes hidden rows/columns)
            JsonSaveOptions jsonOptions = new JsonSaveOptions();

            // Save the workbook as JSON using the configured options
            workbook.Save(outputFile, jsonOptions);

            Console.WriteLine($"Workbook successfully saved to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
