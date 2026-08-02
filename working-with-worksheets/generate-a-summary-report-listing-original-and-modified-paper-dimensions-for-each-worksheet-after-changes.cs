// Title: Aspose.Cells .NET: Generate Paper‑Size Summary Report for All Worksheets
// Description: Loads a workbook, records each worksheet's original PaperWidth and PaperHeight, applies a custom size (e.g., 8.5 × 11 in), captures the new dimensions, writes a "PaperSizeSummary" sheet with the data, and saves the file.
// Keywords: Aspose.Cells | C# | .NET | worksheet paper size | custom paper dimensions | page setup | summary sheet | Excel automation | batch update paper size
// Common Searches: read and change worksheet paper size Aspose.Cells .NET | create summary worksheet with original and new paper dimensions | Aspose.Cells get PaperWidth PaperHeight per sheet | batch set custom paper size for all worksheets
// Developer Intent: Record each sheet’s original paper dimensions, change them to a uniform custom size, and produce a summary worksheet that lists both sets of values.
// Use Cases: Audit current paper settings before printing to verify compliance. | Standardize all worksheets to a specific paper size while keeping a change log. | Provide a printable report of original vs. updated dimensions for documentation.
// AI Prompts: Modify the code to work with centimeters instead of inches for paper dimensions. | Export the original and modified dimensions to a CSV file rather than an Excel sheet. | Apply different custom paper sizes to worksheets based on a configuration dictionary.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsPaperSizeReport
{
    // Simple data holder for dimensions
    // Loads a workbook, records each worksheet's original PaperWidth and PaperHeight, applies a custom size (e.g., 8.5 × 11 in), captures the new dimensions, writes a "PaperSizeSummary" sheet with the data, and saves the file.
    class PaperDimensions
    {
        public string SheetName { get; set; }
        public double OriginalWidth { get; set; }
        public double OriginalHeight { get; set; }
        public double ModifiedWidth { get; set; }
        public double ModifiedHeight { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with actual path)
            Workbook workbook = new Workbook("input.xlsx");

            // List to store dimension info for each worksheet
            List<PaperDimensions> reportData = new List<PaperDimensions>();

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the page setup of the current sheet
                PageSetup pageSetup = sheet.PageSetup;

                // Capture original dimensions (in inches)
                double originalWidth = pageSetup.PaperWidth;
                double originalHeight = pageSetup.PaperHeight;

                // Change the paper size to a custom size (example: 8.5 x 11 inches)
                // This demonstrates modification; you can set any size you need
                pageSetup.CustomPaperSize(8.5, 11.0);

                // After modification, capture the new dimensions
                double modifiedWidth = pageSetup.PaperWidth;
                double modifiedHeight = pageSetup.PaperHeight;

                // Store the data
                reportData.Add(new PaperDimensions
                {
                    SheetName = sheet.Name,
                    OriginalWidth = originalWidth,
                    OriginalHeight = originalHeight,
                    ModifiedWidth = modifiedWidth,
                    ModifiedHeight = modifiedHeight
                });
            }

            // Add a new worksheet to hold the summary report
            Worksheet summarySheet = workbook.Worksheets[workbook.Worksheets.Add()];
            summarySheet.Name = "PaperSizeSummary";

            // Write headers
            summarySheet.Cells[0, 0].PutValue("Worksheet");
            summarySheet.Cells[0, 1].PutValue("Original Width (in)");
            summarySheet.Cells[0, 2].PutValue("Original Height (in)");
            summarySheet.Cells[0, 3].PutValue("Modified Width (in)");
            summarySheet.Cells[0, 4].PutValue("Modified Height (in)");

            // Populate rows with collected data
            for (int i = 0; i < reportData.Count; i++)
            {
                PaperDimensions pd = reportData[i];
                int row = i + 1; // start after header

                summarySheet.Cells[row, 0].PutValue(pd.SheetName);
                summarySheet.Cells[row, 1].PutValue(pd.OriginalWidth);
                summarySheet.Cells[row, 2].PutValue(pd.OriginalHeight);
                summarySheet.Cells[row, 3].PutValue(pd.ModifiedWidth);
                summarySheet.Cells[row, 4].PutValue(pd.ModifiedHeight);
            }

            // Save the workbook with the summary (replace with desired output path)
            workbook.Save("output.xlsx");
        }
    }
}
