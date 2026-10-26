// Title: Generate a summary worksheet that records original and A4 paper dimensions for each sheet using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells how to list original and A4 paper size for each worksheet | Aspose.Cells set worksheet PaperSize to A4 and export dimension report | Create Excel summary sheet with original and modified page setup dimensions using Aspose.Cells | Retrieve PaperWidth and PaperHeight before and after changing to A4 in C# | Automate paper size conversion report for multiple worksheets with Aspose.Cells
// Tags: Aspose.Cells set worksheet paper size to A4 | Aspose.Cells retrieve original paper dimensions | Aspose.Cells generate worksheet dimension summary | Aspose.Cells auto‑fit columns after data population | Aspose.Cells C# page setup dimension report

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPaperDimensionReport
{
    // The C# program loads an Excel workbook, iterates through all worksheets (excluding any existing "Summary" sheet), captures each sheet's original PaperWidth and PaperHeight, changes the page setup to A4 using the PaperSize enum, records the new dimensions, stores the data, adds a new "Summary" worksheet, writes headers and the collected dimension information into columns A‑E, auto‑fits the columns for readability, and saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths to the input and output Excel files
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Ensure the input file exists; create a simple workbook if it does not
                if (!File.Exists(inputPath))
                {
                    var tempWb = new Workbook();
                    tempWb.Worksheets[0].Name = "Sheet1";
                    tempWb.Save(inputPath);
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // List to store original and modified paper dimensions for each worksheet
                var dimensionsInfo = new List<SheetDimensionInfo>();

                // Iterate through all worksheets in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Skip the summary sheet if it already exists (to avoid processing it)
                    if (sheet.Name.Equals("Summary", StringComparison.OrdinalIgnoreCase))
                        continue;

                    // Capture original paper dimensions (in points; 1 point = 1/72 inch)
                    double originalWidth = sheet.PageSetup.PaperWidth;
                    double originalHeight = sheet.PageSetup.PaperHeight;

                    // Modify paper dimensions – set to A4 size using PaperSize enum
                    sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;

                    // Capture modified dimensions after setting PaperSize
                    double modifiedWidth = sheet.PageSetup.PaperWidth;
                    double modifiedHeight = sheet.PageSetup.PaperHeight;

                    // Store the information
                    dimensionsInfo.Add(new SheetDimensionInfo
                    {
                        SheetName = sheet.Name,
                        OriginalWidth = originalWidth,
                        OriginalHeight = originalHeight,
                        ModifiedWidth = modifiedWidth,
                        ModifiedHeight = modifiedHeight
                    });
                }

                // Add a new worksheet for the summary report
                int summaryIndex = workbook.Worksheets.Add();
                Worksheet summarySheet = workbook.Worksheets[summaryIndex];
                summarySheet.Name = "Summary";

                // Write headers
                Cells cells = summarySheet.Cells;
                cells["A1"].PutValue("Worksheet");
                cells["B1"].PutValue("Original Width (pt)");
                cells["C1"].PutValue("Original Height (pt)");
                cells["D1"].PutValue("Modified Width (pt)");
                cells["E1"].PutValue("Modified Height (pt)");

                // Populate the summary data
                int row = 1; // zero‑based index; row 1 is the second row in Excel
                foreach (var info in dimensionsInfo)
                {
                    cells[row, 0].PutValue(info.SheetName);
                    cells[row, 1].PutValue(info.OriginalWidth);
                    cells[row, 2].PutValue(info.OriginalHeight);
                    cells[row, 3].PutValue(info.ModifiedWidth);
                    cells[row, 4].PutValue(info.ModifiedHeight);
                    row++;
                }

                // Auto‑fit columns for better readability
                summarySheet.AutoFitColumns();

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook processed successfully. Output saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Helper class to hold dimension data for a worksheet
        private class SheetDimensionInfo
        {
            public string SheetName { get; set; } = string.Empty;
            public double OriginalWidth { get; set; }
            public double OriginalHeight { get; set; }
            public double ModifiedWidth { get; set; }
            public double ModifiedHeight { get; set; }
        }
    }
}
