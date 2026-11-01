// Title: Loading an Excel workbook with only specific defined names (named ranges) using Aspose.Cells for .NET – limitations and workaround
// Common Searches: Aspose.Cells C# load workbook with only required named ranges | filter defined names when opening an Excel file with Aspose.Cells .NET | exclude unused named ranges during workbook load in Aspose.Cells | calculate formulas after loading workbook with selected defined names in C# | Aspose.Cells LoadOptions to limit defined names in memory
// Tags: Aspose.Cells load workbook with selected defined names | filter named ranges Aspose.Cells .NET | load options exclude unused defined names | calculate formulas after loading workbook Aspose.Cells | working with defined names in Aspose.Cells C# | Excel named ranges loading performance Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example checks for the source XLSX file, loads it using Aspose.Cells LoadOptions (Xlsx format), calculates all formulas, and saves the result. It highlights that the current Aspose.Cells API does not provide a dedicated FilterDefinedNames option, so all defined names are loaded and must be managed manually after loading.
class Program
{
    static void Main()
    {
        // Path to the source workbook
        string sourcePath = "input.xlsx";

        // Path to the output workbook
        string outputPath = "output.xlsx";

        // Ensure the source file exists to prevent FileNotFoundException
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Source file not found: {sourcePath}");
            return;
        }

        try
        {
            // Load the workbook (no specific defined‑name filter available in this API version)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Evaluate all formulas in the workbook
            workbook.CalculateFormula();

            // Save the processed workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
