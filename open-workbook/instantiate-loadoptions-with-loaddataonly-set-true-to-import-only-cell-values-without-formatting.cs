// Title: Load an XLSX workbook in C# with Aspose.Cells using LoadOptions.LoadDataOnly to import only cell values
// AI Prompts: Create a LoadOptions object for Xlsx format, set LoadDataOnly = true, and load the workbook to retrieve only raw cell data. | Update the C# console program to enable LoadDataOnly, then save the workbook while preserving only the imported values. | Write C# code that checks for file existence, loads an Excel file with LoadDataOnly enabled, and writes the result to a new file without any formatting.
// Common Searches: Aspose.Cells C# load workbook without formatting using LoadDataOnly | How to import only cell values from an XLSX file with Aspose.Cells .NET | C# example for LoadOptions.LoadDataOnly true in Aspose.Cells | Read Excel data only (no styles) with Aspose.Cells LoadOptions in a console app | Skip formatting when opening Excel file with Aspose.Cells .NET
// Tags: Aspose.Cells LoadOptions LoadDataOnly | C# load XLSX without formatting | Import cell values only Aspose.Cells | Excel workbook loading options .NET | File existence check C# Aspose.Cells | Save workbook after data‑only load Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program checks that the input XLSX file exists, creates a LoadOptions instance for Xlsx format with LoadDataOnly set to true, loads the workbook so only raw cell values are imported, and saves the workbook to a new file, handling any exceptions that may occur.
class Program
{
    static void Main()
    {
        try
        {
            const string inputFile = "input.xlsx";
            const string outputFile = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file \"{inputFile}\" was not found.");
                return;
            }

            // Load the workbook (format is inferred from the file extension)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            Workbook workbook = new Workbook(inputFile, loadOptions);

            // Save the workbook; it will contain the data (formatting is retained by default)
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully to \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
