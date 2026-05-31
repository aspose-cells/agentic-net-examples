using System;
using System.Diagnostics;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsPerformanceDemo
{
    // Custom globalization settings – overrides boolean display strings
    public class CustomGlobalizationSettings : GlobalizationSettings
    {
        public override string GetBooleanValueString(bool value)
        {
            return value ? "TRUE_CUSTOM" : "FALSE_CUSTOM";
        }
    }

    public class Program
    {
        // Path for the sample workbook
        private const string SampleFile = "sample.xlsx";

        public static void Main()
        {
            // Ensure a sample workbook exists
            CreateSampleWorkbook();

            // Measure load time without custom globalization settings
            TimeSpan defaultLoadTime = MeasureLoadTime(applyCustomSettings: false);
            Console.WriteLine($"Load time without custom globalization settings: {defaultLoadTime.TotalMilliseconds} ms");

            // Measure load time with custom globalization settings
            TimeSpan customLoadTime = MeasureLoadTime(applyCustomSettings: true);
            Console.WriteLine($"Load time with custom globalization settings: {customLoadTime.TotalMilliseconds} ms");
        }

        // Creates a workbook with a sizable amount of data to make loading measurable
        private static void CreateSampleWorkbook()
        {
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate 10,000 rows with sample data
            for (int row = 0; row < 10000; row++)
            {
                cells[row, 0].PutValue($"Item {row}");
                cells[row, 1].PutValue(row);
                cells[row, 2].PutValue(row % 2 == 0);
            }

            // Save the workbook (uses the provided save rule)
            wb.Save(SampleFile);
        }

        // Loads the workbook and optionally applies custom globalization settings,
        // returning the elapsed time for the operation.
        private static TimeSpan MeasureLoadTime(bool applyCustomSettings)
        {
            Stopwatch sw = new Stopwatch();

            // Start timing
            sw.Start();

            // Load the workbook using the standard constructor (provided load rule)
            Workbook wb = new Workbook(SampleFile);

            // Apply custom globalization settings after loading if requested
            if (applyCustomSettings)
            {
                wb.Settings.GlobalizationSettings = new CustomGlobalizationSettings();
            }

            // Stop timing
            sw.Stop();

            // Return elapsed time
            return sw.Elapsed;
        }
    }
}