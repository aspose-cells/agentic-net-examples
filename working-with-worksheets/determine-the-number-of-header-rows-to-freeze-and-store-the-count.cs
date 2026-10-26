// Title: Read the number of frozen header rows in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells get count of frozen rows in a worksheet | how to read FreezePanes.Row value with Aspose.Cells .NET | determine how many header rows are frozen in an Excel file using Aspose.Cells | retrieve frozen pane row index from workbook with Aspose.Cells C#
// Tags: Aspose.Cells FreezePanes.Row retrieval | C# read frozen rows index Excel | worksheet freeze panes inspection Aspose.Cells | determine frozen header rows .NET | reflection based FreezePanes access Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, accesses the first worksheet, and uses reflection to read the FreezePanes.Row property, obtaining the count of frozen header rows, which is then displayed and optionally saved.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook from the input file
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (index 0)
                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve the number of frozen rows (header rows)
                int headerRowsToFreeze = 0;
                try
                {
                    // Use reflection to safely access FreezePanes.Row if the property exists
                    var freezeProp = worksheet.GetType().GetProperty("FreezePanes");
                    if (freezeProp != null)
                    {
                        var freezeObj = freezeProp.GetValue(worksheet);
                        var rowProp = freezeObj?.GetType().GetProperty("Row");
                        if (rowProp != null)
                        {
                            headerRowsToFreeze = (int)rowProp.GetValue(freezeObj);
                        }
                    }
                }
                catch (Exception ex)
                {
                    // If reflection fails, keep default value (0)
                    Console.WriteLine($"Unable to read frozen rows: {ex.Message}");
                }

                Console.WriteLine($"Number of header rows frozen: {headerRowsToFreeze}");

                // Ensure the output directory exists
                try
                {
                    string outputDir = Path.GetDirectoryName(outputPath);
                    if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }

                    // Save the workbook (optional, if any changes were made)
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to: {outputPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                // Handle any runtime exceptions gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
