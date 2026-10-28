// Title: Create a green data bar conditional format for column L to display progress percentages with Aspose.Cells in C#
// Common Searches: how to add a green data bar conditional format to column L using Aspose.Cells C# | Aspose.Cells C# set data bar minimum and maximum values for progress percentages | apply data bar conditional formatting to range L1:L100 in an Excel workbook with Aspose.Cells | visualize percentage progress in Excel using data bars via Aspose.Cells .NET API | C# example for data bar conditional formatting on a specific column with Aspose.Cells
// Tags: Aspose.Cells data bar formatting C# | Excel column L conditional formatting Aspose | set data bar min/max values Aspose.Cells | CellArea range for conditional formatting Aspose | visualize progress percentages with data bars Aspose

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example creates a new workbook, populates column L with random values from 0 to 100 to represent progress percentages, defines the L1:L100 range, adds a conditional formatting collection, inserts a green DataBar condition with minimum 0 and maximum 100, and saves the file as DataBarProgress.xlsx.
    class DataBarConditionalFormattingExample
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Fill column L (index 11) with random progress percentages (0‑100)
                Random rnd = new Random();
                for (int row = 0; row < 100; row++)
                {
                    sheet.Cells[row, 11].PutValue(rnd.Next(0, 101));
                }

                // Define the range L1:L100
                CellArea area = new CellArea
                {
                    StartRow = 0,
                    EndRow = 99,
                    StartColumn = 11,
                    EndColumn = 11
                };

                // Add a conditional formatting collection to the worksheet
                int cfIndex = sheet.ConditionalFormattings.Add();
                FormatConditionCollection cfCollection = sheet.ConditionalFormattings[cfIndex];
                cfCollection.AddArea(area);

                // Add a DataBar condition
                cfCollection.AddCondition(FormatConditionType.DataBar);
                FormatCondition dataBarCondition = cfCollection[0];

                // Configure the DataBar
                DataBar dataBar = dataBarCondition.DataBar;
                dataBar.Color = Color.Green;

                // Set minimum and maximum values (0% – 100%)
                // The default type is numeric, so setting the values is sufficient
                dataBar.MinCfvo.Value = "0";
                dataBar.MaxCfvo.Value = "100";

                // Save the workbook
                string outputPath = "DataBarProgress.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
