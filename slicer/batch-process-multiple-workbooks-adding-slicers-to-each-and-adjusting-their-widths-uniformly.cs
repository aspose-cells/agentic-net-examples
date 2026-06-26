using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class BatchSlicerProcessor
{
    static void Main()
    {
        // Input workbook files to process
        string[] inputFiles = new string[]
        {
            "Workbook1.xlsx",
            "Workbook2.xlsx",
            "Workbook3.xlsx"
        };

        // Folder where processed workbooks will be saved
        string outputFolder = "ProcessedWorkbooks";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Uniform column width for all slicers (in points)
        double uniformColumnWidth = 80.0;

        foreach (string inputPath in inputFiles)
        {
            // Load the workbook from file
            Workbook workbook = new Workbook(inputPath);

            // Work with the first worksheet (adjust as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Proceed only if the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count > 0)
            {
                // Get the first pivot table in the worksheet
                PivotTable pivotTable = worksheet.PivotTables[0];

                // Use the first base field of the pivot table for the slicer
                string baseFieldName = pivotTable.BaseFields[0].Name;

                // Add a slicer anchored at cell "A1"
                int slicerIndex = worksheet.Slicers.Add(pivotTable, "A1", baseFieldName);

                // Retrieve the newly added slicer
                Slicer slicer = worksheet.Slicers[slicerIndex];

                // Set the column width uniformly
                slicer.ColumnWidth = uniformColumnWidth;
            }

            // Determine the output file path
            string outputPath = Path.Combine(outputFolder, Path.GetFileName(inputPath));

            // Save the modified workbook
            workbook.Save(outputPath);

            // Release resources
            workbook.Dispose();
        }
    }
}