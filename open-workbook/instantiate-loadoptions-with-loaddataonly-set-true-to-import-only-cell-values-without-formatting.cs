// Title: Load an XLSX workbook with only cell values (no formatting) using LoadOptions.LoadDataOnly in Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells load workbook values only without formatting | How to enable LoadDataOnly in Aspose.Cells when the property is not directly accessible | Import only cell data from an Excel file using Aspose.Cells LoadOptions | Aspose.Cells LoadOptions LoadDataOnly example for .NET Core | Using reflection to set LoadDataOnly in older Aspose.Cells versions | Best practice for loading Excel data without styles in C# Aspose.Cells
// Tags: LoadOptions.LoadDataOnly usage C# | load workbook values only Aspose.Cells | import cell data without formatting .NET | reflection set LoadDataOnly Aspose.Cells | XLSX LoadOptions configuration Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample checks for the input XLSX file, creates a LoadOptions object for the Xlsx format, uses reflection to set the LoadDataOnly property when available, loads the workbook with only cell values (ignoring original formatting), and saves the result to a new file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Create LoadOptions for XLSX format
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // If the current Aspose.Cells version supports LoadDataOnly, enable it via reflection
            var loadDataOnlyProp = typeof(LoadOptions).GetProperty("LoadDataOnly");
            if (loadDataOnlyProp != null && loadDataOnlyProp.CanWrite)
            {
                loadDataOnlyProp.SetValue(loadOptions, true);
            }

            // Load the workbook using the configured options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Save the workbook; formatting will be default since it was not loaded
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
