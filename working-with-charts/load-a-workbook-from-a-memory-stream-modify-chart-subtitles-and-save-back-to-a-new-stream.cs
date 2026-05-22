using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartSubtitleDemo
{
    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // 1. Create a sample workbook with a chart (uses the create rule)
            // ------------------------------------------------------------
            Workbook sampleWorkbook = new Workbook();
            Worksheet sheet = sampleWorkbook.Worksheets[0];

            // Populate some data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";
            chart.Title.Text = "Sample Chart";
            chart.SubTitle.Text = "Original Subtitle";

            // Save the workbook to a memory stream (uses the save rule)
            MemoryStream inputStream = new MemoryStream();
            sampleWorkbook.Save(inputStream, SaveFormat.Xlsx);
            inputStream.Position = 0; // Reset for reading

            // ------------------------------------------------------------
            // 2. Load the workbook from the memory stream, modify subtitles,
            //    and save back to a new stream.
            // ------------------------------------------------------------
            MemoryStream outputStream = UpdateChartSubtitles(inputStream);

            // ------------------------------------------------------------
            // 3. (Optional) Save the resulting stream to a file to verify.
            // ------------------------------------------------------------
            using (FileStream file = new FileStream("ChartWithUpdatedSubtitle.xlsx", FileMode.Create, FileAccess.Write))
            {
                outputStream.CopyTo(file);
            }

            Console.WriteLine("Chart subtitles updated and workbook saved to 'ChartWithUpdatedSubtitle.xlsx'.");
        }

        /// <summary>
        /// Loads a workbook from the provided stream, updates every chart's subtitle,
        /// and returns a new memory stream containing the modified workbook.
        /// </summary>
        /// <param name="input">MemoryStream containing the original workbook.</param>
        /// <returns>MemoryStream with the modified workbook.</returns>
        private static MemoryStream UpdateChartSubtitles(MemoryStream input)
        {
            // Ensure the stream is positioned at the beginning before loading
            input.Position = 0;

            // Load workbook from stream (uses the Workbook(Stream) constructor rule)
            Workbook wb = new Workbook(input);

            // Iterate through all worksheets and their charts
            foreach (Worksheet ws in wb.Worksheets)
            {
                foreach (Chart ch in ws.Charts)
                {
                    // Modify the subtitle text (uses Chart.SubTitle property)
                    ch.SubTitle.Text = "Updated Subtitle";
                }
            }

            // Save the modified workbook to a new memory stream (uses SaveToStream rule)
            MemoryStream result = wb.SaveToStream();

            // Reset position so the caller can read from the beginning
            result.Position = 0;
            return result;
        }
    }
}