// Title: Export each worksheet of an Excel workbook to separate CSV files with Aspose.Cells for .NET
// Common Searches: how to save every sheet of an Excel file as an individual CSV using Aspose.Cells in C# | Aspose.Cells C# export multiple worksheets to separate CSV files without using Excel Interop | C# code to iterate through workbook worksheets and generate CSV files with Aspose.Cells | save each Excel worksheet as a CSV file programmatically with Aspose.Cells .NET library | convert Excel workbook to raw CSV data files using Aspose.Cells C#
// Tags: Aspose.Cells export worksheet to CSV | C# save Excel sheet as separate CSV file | multiple CSV output per workbook Aspose.Cells | iterate workbook worksheets Aspose.Cells C# | create raw CSV data from Excel using Aspose.Cells | remove default sheet when saving worksheet as CSV Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, loops through each worksheet, creates a temporary workbook containing only that sheet, removes the default empty sheet, and saves the sheet as a CSV file named after the worksheet. It includes file‑existence checks and per‑sheet error handling.
class Program
{
    static void Main()
    {
        // Path to the source Excel workbook
        string sourcePath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Error: The file \"{sourcePath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(sourcePath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    // Create a temporary workbook that contains only the current worksheet
                    Workbook sheetWorkbook = new Workbook();

                    // Add a copy of the current sheet by name
                    sheetWorkbook.Worksheets.AddCopy(sheet.Name);

                    // Remove the default empty sheet that is created with a new workbook
                    sheetWorkbook.Worksheets.RemoveAt(0);

                    // Create an output file name based on the worksheet name
                    string outputPath = $"{sheet.Name}.csv";

                    // Save the worksheet as a CSV file (comma delimiter is default)
                    sheetWorkbook.Save(outputPath, SaveFormat.Csv);
                    Console.WriteLine($"Worksheet \"{sheet.Name}\" saved to \"{outputPath}\".");
                }
                catch (Exception ex)
                {
                    // Handle errors for individual worksheets without stopping the whole process
                    Console.WriteLine($"Failed to save worksheet \"{sheet.Name}\": {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display an error message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
