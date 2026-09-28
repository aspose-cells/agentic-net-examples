// Title: Apply Compact Layout to the First PivotTable in an Existing Excel Workbook Using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an existing XLSX file with Aspose.Cells, locates the first PivotTable on the first worksheet, and switches its layout to Compact using the ShowInCompactForm method. | Demonstrate how to verify the presence of PivotTables, apply the Compact layout, and save the modified workbook to a new file with Aspose.Cells. | Create a .NET example that handles missing input files, creates the output directory if needed, updates a PivotTable's layout to Compact, and then saves the workbook.
// Common Searches: Aspose.Cells C# how to set pivot table to compact layout | ShowInCompactForm method example for existing workbook | Change pivot table layout to compact using Aspose.Cells .NET | Update first pivot table layout in Excel file with Aspose.Cells | C# code to apply compact form to pivot table and save workbook
// Tags: apply compact layout pivot table Aspose.Cells C# | ShowInCompactForm method usage | modify pivot table layout in existing workbook | load and save workbook after pivot table changes Aspose.Cells | first worksheet pivot table compact form

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExample
{
    // The example loads an existing XLSX file, checks for a PivotTable on the first worksheet, applies the Compact layout via ShowInCompactForm, and saves the updated workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Get the collection of pivot tables on the worksheet
                PivotTableCollection pivotTables = worksheet.PivotTables;

                if (pivotTables != null && pivotTables.Count > 0)
                {
                    // Select the first pivot table
                    PivotTable pivotTable = pivotTables[0];

                    // Apply Compact layout using the correct API method (parameterless)
                    pivotTable.ShowInCompactForm();
                }
                else
                {
                    Console.WriteLine("No pivot tables found in the worksheet.");
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook with the updated pivot table layout
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
