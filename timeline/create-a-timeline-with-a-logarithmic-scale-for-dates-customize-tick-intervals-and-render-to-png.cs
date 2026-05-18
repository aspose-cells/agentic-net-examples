using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsTimelineLogChart
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (or load a template if it exists)
                string templatePath = "Template.xlsx";
                Workbook workbook = File.Exists(templatePath) ? new Workbook(templatePath) : new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data: dates in column A and numeric values in column B
                cells["A1"].PutValue("Date");
                cells["B1"].PutValue("Value");
                DateTime start = new DateTime(2023, 1, 1);
                for (int i = 0; i < 10; i++)
                {
                    // Exponential dates to illustrate logarithmic scaling
                    cells[i + 2, 0].PutValue(start.AddDays(Math.Pow(2, i)));
                    // Sample numeric value
                    cells[i + 2, 1].PutValue(i + 1);
                }

                // Save the workbook
                string outputPath = "TimelineLogChart.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}