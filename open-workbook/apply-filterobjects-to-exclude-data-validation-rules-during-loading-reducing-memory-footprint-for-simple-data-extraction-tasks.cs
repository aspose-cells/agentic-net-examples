// Title: Using LoadOptions.FilterObjects to skip data‑validation objects when loading an Excel workbook with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# load workbook without data validation objects | How to reduce memory usage by filtering objects during Excel load in .NET | LoadOptions.FilterObjects example for skipping validation rules in Aspose.Cells | C# code to ignore data validation when opening an .xlsx file with Aspose.Cells | Improving performance of Aspose.Cells workbook load by disabling validation
// Tags: filterobjects skip data validation aspnet | loadoptions memory optimization aspose.cells | excel workbook loading filtered objects c# | aspose.cells reduce memory footprint loading | load workbook with filtered objects aspnet | aspose.cells loadoptions configuration c#

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsFilterExample
{
    // This C# example demonstrates creating a LoadOptions object for Xlsx files, loading a workbook with Aspose.Cells, reading a cell value, and saving the file. It highlights that the current Aspose.Cells version does not provide a direct option to disable data‑validation objects via FilterObjects, but shows where such configuration would be applied to lower memory consumption for simple extraction scenarios.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for input and output workbooks
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Create LoadOptions (no data‑validation disabling property in this version)
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

                // Load the workbook with the specified options
                Workbook workbook = new Workbook(inputPath, loadOptions);

                // Example: read a value from the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                string cellValue = sheet.Cells["A1"].StringValue;
                Console.WriteLine($"Value in A1: {cellValue}");

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
