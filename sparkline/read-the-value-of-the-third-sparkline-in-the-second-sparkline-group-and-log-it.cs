using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSparklineDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for two sparkline groups
                // Group 1 data (row 0)
                sheet.Cells["A1"].PutValue(5);
                sheet.Cells["B1"].PutValue(3);
                sheet.Cells["C1"].PutValue(8);
                sheet.Cells["D1"].PutValue(2);
                // Group 2 data (row 1)
                sheet.Cells["A2"].PutValue(7);
                sheet.Cells["B2"].PutValue(1);
                sheet.Cells["C2"].PutValue(4);
                sheet.Cells["D2"].PutValue(6);

                // Define locations where sparklines will be placed
                // First group will be placed in column E (index 4) row 0
                CellArea locationGroup1 = new CellArea { StartRow = 0, EndRow = 0, StartColumn = 4, EndColumn = 4 };
                // Second group will be placed in column F (index 5) row 1
                CellArea locationGroup2 = new CellArea { StartRow = 1, EndRow = 1, StartColumn = 5, EndColumn = 5 };

                // Add first sparkline group (shown for completeness)
                int groupIdx1 = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, locationGroup1);
                SparklineGroup group1 = sheet.SparklineGroups[groupIdx1];
                // Add three sparklines to the first group (same location)
                group1.Sparklines.Add("A1:D1", 0, 4);
                group1.Sparklines.Add("A1:D1", 0, 4);
                group1.Sparklines.Add("A1:D1", 0, 4);

                // Add second sparkline group – this is the group we will read from
                int groupIdx2 = sheet.SparklineGroups.Add(SparklineType.Line, "A2:D2", false, locationGroup2);
                SparklineGroup group2 = sheet.SparklineGroups[groupIdx2];
                // Add three sparklines to the second group
                group2.Sparklines.Add("A2:D2", 1, 5); // sparkline 0
                group2.Sparklines.Add("A2:D2", 1, 5); // sparkline 1
                group2.Sparklines.Add("A2:D2", 1, 5); // sparkline 2

                // Access the third sparkline (index 2) in the second sparkline group
                Sparkline thirdSparkline = sheet.SparklineGroups[groupIdx2].Sparklines[2];

                // Log its DataRange – this represents the source data for the sparkline
                Console.WriteLine("Third sparkline DataRange: " + thirdSparkline.DataRange);

                // Demonstrate extracting the first cell value from the DataRange
                string range = thirdSparkline.DataRange; // e.g., "Sheet1!A2:D2"
                string address = range.Contains("!") ? range.Split('!')[1] : range;
                Cell firstCell = sheet.Cells[address];
                Console.WriteLine("First cell value in the sparkline range: " + firstCell.Value);

                // Save the workbook (output file)
                string outputPath = "SparklineDemoOutput.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}