// Title: How to validate that a PivotTable contains at least one data field before refreshing and calculating with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# check if a PivotTable contains any DataFields before calling RefreshData | How to verify PivotTable DataFields count using Aspose.Cells in .NET | Prevent InvalidOperationException when calculating a PivotTable without data fields in C# | Refresh and calculate a PivotTable only when it has at least one data field with Aspose.Cells | Validate all PivotTables in a workbook for data fields before processing in .NET
// Tags: Aspose.Cells PivotTable data field check | C# ensure PivotTable has data fields | RefreshData after PivotTable validation | CalculateData with data field verification | PivotTable collection validation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable and PivotTableCollection

namespace PivotTableValidation
{
    // The example loads a workbook, iterates through each PivotTable on the first worksheet, throws an exception if any PivotTable lacks data fields, then safely refreshes the cache and recalculates the PivotTable before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input workbook exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                    throw new FileNotFoundException($"The input file \"{inputPath}\" was not found.");

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index as needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Get the collection of PivotTables on the worksheet
                PivotTableCollection pivotTables = worksheet.PivotTables;

                // Ensure there is at least one PivotTable to validate
                if (pivotTables.Count == 0)
                    throw new InvalidOperationException("No PivotTable found on the worksheet.");

                // Validate each PivotTable for at least one data field and refresh it
                foreach (PivotTable pivot in pivotTables)
                {
                    // Check if the PivotTable contains any data fields
                    if (pivot.DataFields.Count == 0)
                        throw new InvalidOperationException($"PivotTable \"{pivot.Name}\" must contain at least one data field before calculations.");

                    // Refresh the underlying cache and recalculate the PivotTable
                    pivot.RefreshData();      // Refreshes the cache associated with the PivotTable
                    pivot.CalculateData();    // Recalculates the PivotTable
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook after successful validation and calculations
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Log the exception details for troubleshooting
                Console.Error.WriteLine($"Error: {ex.Message}");
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
