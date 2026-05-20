using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Slicers;

namespace BatchSlicerProcessor
{
    class Program
    {
        static void Main()
        {
            // Folder containing source workbooks
            string sourceFolder = @"C:\InputWorkbooks";
            // Folder where processed workbooks will be saved
            string outputFolder = @"C:\OutputWorkbooks";

            // Ensure output folder exists
            Directory.CreateDirectory(outputFolder);

            // Uniform slicer column width (in points)
            double uniformColumnWidth = 100.0;

            // Process each .xlsx file in the source folder
            foreach (string inputPath in Directory.GetFiles(sourceFolder, "*.xlsx"))
            {
                // Load existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Work with the first worksheet (adjust as needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // Create a simple table if the worksheet has no tables
                // (rows 0-4, columns 0-1) – this ensures a data source for the slicer
                // -------------------------------------------------
                int tableIndex = worksheet.ListObjects.Add(0, 0, 4, 1, true);
                ListObject table = worksheet.ListObjects[tableIndex];

                // Add a slicer for the first column of the table at cell E1
                int slicerIndex = worksheet.Slicers.Add(table, 0, "E1");
                Slicer slicer = worksheet.Slicers[slicerIndex];

                // Apply the uniform column width to the slicer
                slicer.ColumnWidth = uniformColumnWidth;

                // Save the modified workbook to the output folder
                string fileName = Path.GetFileName(inputPath);
                string outputPath = Path.Combine(outputFolder, fileName);
                workbook.Save(outputPath);
            }
        }
    }
}