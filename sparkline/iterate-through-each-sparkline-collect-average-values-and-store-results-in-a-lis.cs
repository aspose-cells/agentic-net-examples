using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeRange = Aspose.Cells.Range;

namespace SparklineAverageDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for two rows (two sparklines)
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(2);
            sheet.Cells["C1"].PutValue(1);
            sheet.Cells["D1"].PutValue(3);

            sheet.Cells["A2"].PutValue(8);
            sheet.Cells["B2"].PutValue(4);
            sheet.Cells["C2"].PutValue(6);
            sheet.Cells["D2"].PutValue(2);

            // Define where the sparklines will be placed (column E)
            CellArea location = new CellArea
            {
                StartRow = 0,
                EndRow = 1,
                StartColumn = 4,
                EndColumn = 4
            };

            // Add a sparkline group for the two rows
            int groupIdx = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D2", false, location);
            SparklineGroup group = sheet.SparklineGroups[groupIdx];

            // Add individual sparklines for each row
            group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4); // first row
            group.Sparklines.Add(sheet.Name + "!A2:D2", 1, 4); // second row

            // List to hold average values of each sparkline
            List<double> sparklineAverages = new List<double>();

            // Iterate through each sparkline in the group
            foreach (Sparkline sparkline in group.Sparklines)
            {
                // DataRange may contain sheet name (e.g., "Sheet1!A1:D1")
                string range = sparkline.DataRange;
                string sheetName = sheet.Name; // default to current sheet
                string address = range;

                // If a sheet name is present, split it
                if (range.Contains("!"))
                {
                    string[] parts = range.Split('!');
                    sheetName = parts[0].Trim('\''); // remove possible quotes
                    address = parts[1];
                }

                // Get the worksheet that contains the data range
                Worksheet dataSheet = workbook.Worksheets[sheetName];

                // Obtain the range object
                AsposeRange dataRange = dataSheet.Cells.CreateRange(address);

                // Compute the average of all numeric cells in the range
                double sum = 0;
                int count = 0;
                foreach (Cell cell in dataRange)
                {
                    if (cell.Value != null && double.TryParse(cell.Value.ToString(), out double val))
                    {
                        sum += val;
                        count++;
                    }
                }

                double average = count > 0 ? sum / count : 0;
                sparklineAverages.Add(average);
            }

            // Output the averages (for demonstration)
            for (int i = 0; i < sparklineAverages.Count; i++)
            {
                Console.WriteLine($"Sparkline {i} average: {sparklineAverages[i]}");
            }

            // Save the workbook
            workbook.Save("SparklineAverages.xlsx");
        }
    }
}