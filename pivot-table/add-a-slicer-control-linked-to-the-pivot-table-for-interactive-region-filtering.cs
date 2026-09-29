// Title: Create a Region slicer linked to the first pivot table in an Excel file with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an existing .xlsx workbook, finds the 'Region' row field in the first pivot table, creates a slicer for that field at cell A5, sets the slicer name to 'RegionSlicer' and custom height/width, then saves the workbook. | Show how to obtain the BaseFieldIndex of a pivot table row field and use the Worksheet.Slicers.Add method to attach a slicer to a pivot table in Aspose.Cells. | Demonstrate checking for the input workbook, ensuring the output directory exists, and handling exceptions while adding a slicer to a pivot table with Aspose.Cells.
// Common Searches: Aspose.Cells C# create region slicer for pivot table | Link Excel slicer to pivot table using Aspose.Cells .NET | C# example of slicer placement at cell A5 with Aspose.Cells | Retrieve BaseFieldIndex of pivot table row field in Aspose.Cells | Save workbook after modifying slicer with Aspose.Cells C#
// Tags: slicer creation for pivot table Aspose.Cells | region field slicer Excel .NET | slicers API usage Aspose.Cells | pivot table basefieldindex lookup C# | output directory handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsExample
{
    // // Loads an existing workbook, locates the 'Region' row field in the first pivot table, adds a slicer linked to that field at cell A5, names it 'RegionSlicer', sets its size, ensures the output folder exists, and saves the modified file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure a pivot table exists on the worksheet
                if (worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No pivot tables found on the first worksheet.");
                    return;
                }

                // Retrieve the first pivot table
                PivotTable pivotTable = worksheet.PivotTables[0];

                // Find the index of the "Region" field in the source data (BaseField)
                int fieldIndex = -1;
                foreach (PivotField pf in pivotTable.RowFields)
                {
                    if (pf.Name.Equals("Region", StringComparison.OrdinalIgnoreCase))
                    {
                        // Use BaseFieldIndex to get the source data column index
                        fieldIndex = pf.BaseFieldIndex;
                        break;
                    }
                }

                if (fieldIndex == -1)
                {
                    Console.WriteLine("Region field not found in the pivot table.");
                    return;
                }

                // Add a slicer linked to the pivot table for the "Region" field
                // Overload: Add(PivotTable, fieldIndex, firstRow, firstColumnName)
                int slicerIndex = worksheet.Slicers.Add(pivotTable, fieldIndex, 5, "A");

                // Configure the created slicer
                Slicer slicer = worksheet.Slicers[slicerIndex];
                slicer.Name = "RegionSlicer";
                slicer.Height = 150; // Height in points
                slicer.Width = 200;  // Width in points

                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook with the new slicer control
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
