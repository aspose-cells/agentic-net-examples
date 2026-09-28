// Title: Add slicers with uniform column width to every pivot table in multiple Excel workbooks using Aspose.Cells for .NET
// AI Prompts: Insert a slicer for the first base field of each pivot table in every worksheet of all .xlsx files in a folder, then set the slicer’s column width to a given point size with Aspose.Cells. | Iterate through a directory of workbooks, add slicers to pivot tables, apply a consistent column width, and save the updated files to a separate output directory using C#.
// Common Searches: C# how to programmatically add a slicer to each pivot table in a batch of Excel files with Aspose.Cells | set the same slicer column width for multiple workbooks using Aspose.Cells .NET | automate slicer creation for all worksheets in a folder of .xlsx files | Aspose.Cells batch processing example for adding slicers and adjusting width | process multiple Excel workbooks to add slicers based on first pivot field in C#
// Tags: batch add slicers Aspose.Cells .NET | uniform slicer column width C# | process multiple .xlsx workbooks Aspose.Cells | pivot table slicer automation Aspose.Cells | slicer column width setting Aspose.Cells API

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

// The program scans a specified input folder for .xlsx files, loads each workbook with Aspose.Cells, iterates through all worksheets, adds a slicer based on the first base field of the first pivot table, sets the slicer’s column width to a uniform value, and saves the modified workbook to an output folder.
class BatchSlicerProcessor
{
    static void Main()
    {
        // Folder containing source workbooks
        string inputFolder = "InputWorkbooks";
        // Folder where processed workbooks will be saved
        string outputFolder = "OutputWorkbooks";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Verify input folder exists; if not, create it and exit (no files to process)
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder \"{inputFolder}\" does not exist. Creating it now.");
            Directory.CreateDirectory(inputFolder);
            Console.WriteLine("Place .xlsx files in the input folder and rerun the program.");
            return;
        }

        // Desired uniform column width for all slicers (in points)
        double uniformWidth = 80.0;

        // Process each .xlsx file in the input folder
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx"))
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                // Load the workbook
                using (Workbook workbook = new Workbook(filePath))
                {
                    // Iterate through all worksheets in the workbook
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        // Proceed only if the worksheet contains at least one pivot table
                        if (sheet.PivotTables.Count > 0)
                        {
                            // Use the first pivot table as the data source for the slicer
                            PivotTable pivot = sheet.PivotTables[0];

                            // Determine the field name to base the slicer on (first base field)
                            if (pivot.BaseFields.Count == 0)
                                continue; // No base fields to create a slicer

                            string baseFieldName = pivot.BaseFields[0].Name;

                            // Add a slicer at cell A1
                            int slicerIdx = sheet.Slicers.Add(pivot, "A1", baseFieldName);

                            // Retrieve the newly added slicer
                            Slicer slicer = sheet.Slicers[slicerIdx];

                            // Set the column width uniformly
                            slicer.ColumnWidth = uniformWidth;
                        }
                    }

                    // Save the modified workbook to the output folder
                    string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
                    workbook.Save(outputPath);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file \"{filePath}\": {ex.Message}");
            }
        }
    }
}
