using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartTitleFontExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each chart in the current worksheet
                foreach (Chart chart in sheet.Charts)
                {
                    // Ensure the chart has a title and it is visible
                    if (chart.Title != null)
                    {
                        chart.Title.IsVisible = true;

                        // Set the title font to Arial, size 12
                        chart.Title.Font.Name = "Arial";
                        chart.Title.Font.Size = 12;
                    }
                }
            }

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("output.xlsx");
        }
    }
}