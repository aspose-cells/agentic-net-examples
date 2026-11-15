// Title: Generate a dimensions summary worksheet listing original and modified paper size for each sheet with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# get original worksheet paper width and height before changing page setup | How to create a summary sheet that records paper dimensions of all worksheets in a workbook using Aspose.Cells | C# set worksheet paper size to A4 and capture previous dimensions with Aspose.Cells | Generate a report of page setup sizes for each Excel sheet and save it to a new worksheet in .NET | Aspose.Cells example to list original and modified page dimensions for every worksheet
// Tags: aspnetcells capture worksheet paper dimensions | aspnetcells set worksheet paper size to A4 | aspnetcells create dimensions summary worksheet | aspnetcells export sheet size report to new sheet | aspnetcells record original and modified page setup sizes

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The code loads an existing workbook (or creates a new one), iterates through all worksheets to record each sheet's original PaperWidth and PaperHeight, sets the paper size to A4, captures the new dimensions, stores the data, adds a new worksheet named "Dimensions Summary" with headers and the collected size information, and saves the workbook.
class Program
{
    static void Main()
    {
        // Paths for input and output workbooks
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook with one worksheet
                Console.WriteLine($"Input file not found. A new workbook has been created.");
            }

            // List to store original and modified paper dimensions for each worksheet
            List<WorksheetInfo> dimensions = new List<WorksheetInfo>();

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Capture original paper dimensions (in points)
                double originalWidth = sheet.PageSetup.PaperWidth;
                double originalHeight = sheet.PageSetup.PaperHeight;

                // Set paper size to A4 (Aspose.Cells uses PaperSizeType enum)
                sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;

                // Capture modified dimensions after setting the paper size
                double modifiedWidth = sheet.PageSetup.PaperWidth;
                double modifiedHeight = sheet.PageSetup.PaperHeight;

                // Store the dimensions for reporting
                dimensions.Add(new WorksheetInfo
                {
                    Name = sheet.Name,
                    OriginalWidth = originalWidth,
                    OriginalHeight = originalHeight,
                    ModifiedWidth = modifiedWidth,
                    ModifiedHeight = modifiedHeight
                });
            }

            // Add a new worksheet to hold the summary report
            int summaryIndex = workbook.Worksheets.Add();
            Worksheet summarySheet = workbook.Worksheets[summaryIndex];
            summarySheet.Name = "Dimensions Summary";

            // Write header row
            summarySheet.Cells["A1"].PutValue("Worksheet");
            summarySheet.Cells["B1"].PutValue("Original Width (pt)");
            summarySheet.Cells["C1"].PutValue("Original Height (pt)");
            summarySheet.Cells["D1"].PutValue("Modified Width (pt)");
            summarySheet.Cells["E1"].PutValue("Modified Height (pt)");

            // Populate the summary data (row index is zero‑based)
            int rowIndex = 1;
            foreach (var info in dimensions)
            {
                summarySheet.Cells[rowIndex, 0].PutValue(info.Name);
                summarySheet.Cells[rowIndex, 1].PutValue(info.OriginalWidth);
                summarySheet.Cells[rowIndex, 2].PutValue(info.OriginalHeight);
                summarySheet.Cells[rowIndex, 3].PutValue(info.ModifiedWidth);
                summarySheet.Cells[rowIndex, 4].PutValue(info.ModifiedHeight);
                rowIndex++;
            }

            // Save the modified workbook with the summary report
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper class to hold dimension information for each worksheet
    class WorksheetInfo
    {
        public string Name { get; set; } = string.Empty;
        public double OriginalWidth { get; set; }
        public double OriginalHeight { get; set; }
        public double ModifiedWidth { get; set; }
        public double ModifiedHeight { get; set; }
    }
}
