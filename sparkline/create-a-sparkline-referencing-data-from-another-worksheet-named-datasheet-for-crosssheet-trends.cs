// Title: Create a line sparkline on a worksheet that references a range in another worksheet with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code using Aspose.Cells to add a line sparkline to cell E1 on Sheet1 that reads data from DataSheet!A1:D1. | Show how to define a CellArea for sparkline placement and configure the SparklineGroup to plot by row. | Provide the steps to save the workbook as an .xlsx file after creating the cross‑sheet sparkline.
// Common Searches: Aspose.Cells C# add sparkline that uses data from a different worksheet | How to create a line sparkline referencing DataSheet range in Aspose.Cells .NET | Cross‑sheet sparkline example with Aspose.Cells for .NET | Set sparkline location cell E1 using Aspose.Cells C# | Programmatically save workbook after adding sparkline in Aspose.Cells
// Tags: add line sparkline cross‑sheet Aspose.Cells C# | define sparkline CellArea placement Aspose.Cells | save workbook as xlsx after sparkline Aspose.Cells | sparkline group plot by row .NET | reference external worksheet range for sparkline Aspose

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineCrossSheetDemo
{
    // Demonstrates creating a workbook, adding a DataSheet with sample values, inserting a line sparkline on a separate sheet that references DataSheet!A1:D1, positioning it in cell E1, and saving the file as CrossSheetSparkline.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a worksheet that will hold the source data and name it "DataSheet"
            Worksheet dataSheet = workbook.Worksheets[workbook.Worksheets.Add()];
            dataSheet.Name = "DataSheet";

            // Populate sample data in DataSheet (range A1:D1)
            dataSheet.Cells["A1"].PutValue(5);
            dataSheet.Cells["B1"].PutValue(2);
            dataSheet.Cells["C1"].PutValue(8);
            dataSheet.Cells["D1"].PutValue(3);

            // Use the first worksheet (index 0) to place the sparkline
            Worksheet sparklineSheet = workbook.Worksheets[0];
            sparklineSheet.Name = "SparklineSheet";

            // Define where the sparkline will be placed (cell E1)
            CellArea location = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 4, // Column E (0‑based index)
                EndColumn = 4
            };

            // Add a sparkline group that references the data on "DataSheet"
            int groupIndex = sparklineSheet.SparklineGroups.Add(
                SparklineType.Line,          // Sparkline type
                "DataSheet!A1:D1",           // Cross‑sheet data range
                false,                       // Plot by row (horizontal)
                location);                   // Where the sparkline will appear

            // Retrieve the created group (optional, for further customization)
            SparklineGroup group = sparklineSheet.SparklineGroups[groupIndex];

            // The sparkline is already created by the Add method above.
            // Additional customization can be done here if needed, e.g.:
            // group.ShowHighPoint = true;
            // group.ShowLowPoint = true;

            // Save the workbook to a file
            workbook.Save("CrossSheetSparkline.xlsx");
        }
    }
}
