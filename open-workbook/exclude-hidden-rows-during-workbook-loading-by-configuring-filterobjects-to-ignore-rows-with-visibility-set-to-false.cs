// Title: How to exclude hidden rows from an XLSX workbook when loading with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, iterates every worksheet, deletes rows where Row.IsHidden is true, and saves the result to a new file. | Create a reusable method that takes input and output paths, validates the source file, loads the workbook using LoadOptions, removes hidden rows in reverse order to preserve indices, and returns a success flag. | Write a script that logs each hidden row removed per worksheet while processing an Excel workbook with Aspose.Cells, then writes the cleaned workbook to a specified destination.
// Common Searches: aspnet load excel file ignore hidden rows aspose.cells | c# aspose.cells remove hidden rows before saving workbook | how to filter out hidden rows when opening an xlsx with Aspose.Cells | exclude hidden rows during workbook load using Aspose.Cells .NET | remove hidden rows from Excel using Aspose.Cells LoadOptions
// Tags: remove hidden rows Aspose.Cells | Aspose.Cells load workbook without hidden rows | C# delete hidden rows Excel | Row.IsHidden Aspose.Cells | LoadOptions Xlsx Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example checks that the source XLSX file exists, loads it with Aspose.Cells LoadOptions, iterates each worksheet, removes rows flagged as hidden by evaluating Row.IsHidden (processing rows in reverse order to avoid index shifts), and saves the cleaned workbook to a new file.
class Program
{
    static void Main()
    {
        // Paths for the source and destination files
        string inputFile = "input.xlsx";
        string outputFile = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file not found: {inputFile}");
                return;
            }

            // LoadOptions: configure the loader (no FilterObjects property in Aspose.Cells)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // Load the workbook with the configured options
            Workbook workbook = new Workbook(inputFile, loadOptions);

            // Remove any rows that are hidden after loading
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate from the last row to the first to safely remove rows
                for (int rowIndex = sheet.Cells.MaxDataRow; rowIndex >= sheet.Cells.MinDataRow; rowIndex--)
                {
                    Row row = sheet.Cells.Rows[rowIndex];
                    if (row.IsHidden)
                    {
                        // Remove the hidden row by index
                        sheet.Cells.Rows.RemoveAt(rowIndex);
                    }
                }
            }

            // Save the cleaned workbook
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully to {outputFile}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
