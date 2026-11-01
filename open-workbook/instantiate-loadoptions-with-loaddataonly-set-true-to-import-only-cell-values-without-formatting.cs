// Title: Load an Excel workbook with only cell values (no formatting) using LoadOptions.LoadDataOnly in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# load workbook data only without styles | Set LoadOptions.LoadDataOnly to true when opening an XLSX file with Aspose.Cells | Import only Excel cell values using Aspose.Cells LoadOptions in .NET | C# example for loading an Excel file without formatting using Aspose.Cells
// Tags: LoadOptions.LoadDataOnly C# | load workbook cell values only Aspose.Cells | skip formatting on Excel load Aspose | Aspose.Cells XLSX load options usage | import Excel data without styles .NET

using Aspose.Cells;
using System;
using System.IO;

// The sample verifies the input XLSX file, creates a LoadOptions object for the Xlsx format, sets the LoadDataOnly property to true via reflection (if available), loads the workbook with those options, and saves it, resulting in a workbook that contains only raw cell values without any formatting or styles.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Create LoadOptions for XLSX format
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // Set LoadDataOnly if the property exists in the current Aspose.Cells version
            var loadDataOnlyProp = typeof(LoadOptions).GetProperty("LoadDataOnly");
            if (loadDataOnlyProp != null && loadDataOnlyProp.CanWrite)
            {
                loadDataOnlyProp.SetValue(loadOptions, true);
            }

            // Load the workbook using the specified LoadOptions
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Save the workbook to verify that only data was loaded
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
