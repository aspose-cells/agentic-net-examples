// Title: Aspose.Cells .NET example: Generate a worksheet‑wise report of original and custom paper dimensions
// Description: C# code that creates a workbook with three sheets, reads each sheet's built‑in PaperWidth and PaperHeight, changes every sheet to a 6 × 8 in custom size, records the new dimensions, and writes a "Summary" worksheet listing the sheet name, original width/height, and modified width/height before saving as PaperSizeReport.xlsx.
// Keywords: Aspose.Cells | C# | Worksheet paper size | PaperWidth | PaperHeight | CustomPaperSize | PageSetup | Excel summary report | Workbook automation | Aspose.Cells example
// Common Searches: how to get worksheet paper dimensions with Aspose.Cells | Aspose.Cells change all sheets to custom paper size | create summary sheet with original and new page setup values Aspose.Cells | C# Aspose.Cells report paper width and height per worksheet | Aspose.Cells skip summary worksheet during iteration
// Developer Intent: Log each worksheet's original page dimensions, apply a uniform 6 × 8 in custom paper size, and output both sets of values to a new "Summary" sheet in the same workbook.
// Use Cases: Audit existing worksheets before enforcing a standard paper size across a workbook. | Generate compliance documentation that shows before‑and‑after page dimensions for each sheet. | Automate batch processing to standardize paper size while preserving a change log for downstream review.
// AI Prompts: Write C# Aspose.Cells code that iterates through all worksheets, captures PaperWidth and PaperHeight, sets a custom 6x8‑inch paper size, and adds a "Summary" sheet with the collected data. | Show how to round the original and modified dimensions to two decimal places and save the workbook as PaperSizeReport.xlsx. | Explain how to exclude a worksheet named "Summary" from the iteration to prevent it from being modified.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsPaperSizeReport
{
    // C# code that creates a workbook with three sheets, reads each sheet's built‑in PaperWidth and PaperHeight, changes every sheet to a 6 × 8 in custom size, records the new dimensions, and writes a "Summary" worksheet listing the sheet name, original width/height, and modified width/height before saving as PaperSizeReport.xlsx.
    class Program
    {
        // Simple data holder for dimensions
        class PaperDimensions
        {
            public string SheetName { get; set; }
            public double OriginalWidth { get; set; }
            public double OriginalHeight { get; set; }
            public double ModifiedWidth { get; set; }
            public double ModifiedHeight { get; set; }
        }

        static void Main()
        {
            // -------------------------------------------------
            // 1. Create a new workbook and add sample worksheets
            // -------------------------------------------------
            Workbook workbook = new Workbook(); // create workbook
            // Ensure we have at least three worksheets for demonstration
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
            Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");

            // Set different initial paper sizes for each sheet
            sheet1.PageSetup.PaperSize = PaperSizeType.PaperLetter;          // 8.5 x 11 in
            sheet2.PageSetup.PaperSize = PaperSizeType.PaperLegal;          // 8.5 x 14 in
            sheet3.PageSetup.PaperSize = PaperSizeType.PaperA5;             // 148 x 210 mm (~5.83 x 8.27 in)

            // -------------------------------------------------
            // 2. Capture original dimensions, modify paper size,
            //    and capture modified dimensions
            // -------------------------------------------------
            List<PaperDimensions> reportData = new List<PaperDimensions>();

            // Iterate over all worksheets except the summary sheet (which does not exist yet)
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // Skip any sheet that will be used for the report later
                if (ws.Name.Equals("Summary", StringComparison.OrdinalIgnoreCase))
                    continue;

                PageSetup ps = ws.PageSetup;

                // Record original dimensions (in inches)
                double originalWidth = ps.PaperWidth;
                double originalHeight = ps.PaperHeight;

                // Modify the paper size:
                // For demonstration, set a custom size of 6 inches x 8 inches for every sheet
                ps.CustomPaperSize(6.0, 8.0);

                // Record modified dimensions
                double modifiedWidth = ps.PaperWidth;
                double modifiedHeight = ps.PaperHeight;

                // Store the data
                reportData.Add(new PaperDimensions
                {
                    SheetName = ws.Name,
                    OriginalWidth = originalWidth,
                    OriginalHeight = originalHeight,
                    ModifiedWidth = modifiedWidth,
                    ModifiedHeight = modifiedHeight
                });
            }

            // -------------------------------------------------
            // 3. Create a summary worksheet and populate it
            // -------------------------------------------------
            Worksheet summarySheet = workbook.Worksheets.Add("Summary");
            Cells cells = summarySheet.Cells;

            // Write header row
            cells["A1"].PutValue("Worksheet");
            cells["B1"].PutValue("Original Width (in)");
            cells["C1"].PutValue("Original Height (in)");
            cells["D1"].PutValue("Modified Width (in)");
            cells["E1"].PutValue("Modified Height (in)");

            // Populate rows
            int rowIndex = 1; // zero‑based index; row 1 is the second row (after header)
            foreach (var data in reportData)
            {
                cells[rowIndex, 0].PutValue(data.SheetName);
                cells[rowIndex, 1].PutValue(Math.Round(data.OriginalWidth, 2));
                cells[rowIndex, 2].PutValue(Math.Round(data.OriginalHeight, 2));
                cells[rowIndex, 3].PutValue(Math.Round(data.ModifiedWidth, 2));
                cells[rowIndex, 4].PutValue(Math.Round(data.ModifiedHeight, 2));
                rowIndex++;
            }

            // -------------------------------------------------
            // 4. Save the workbook
            // -------------------------------------------------
            workbook.Save("PaperSizeReport.xlsx");
        }
    }
}
