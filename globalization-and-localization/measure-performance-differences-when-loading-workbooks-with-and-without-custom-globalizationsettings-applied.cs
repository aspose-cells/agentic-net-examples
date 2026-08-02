using System;
using System.Diagnostics;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsPerformanceDemo
{
    // Custom globalization settings – overrides a few methods for demonstration.
    public class CustomGlobalizationSettings : GlobalizationSettings
    {
        public override string GetBooleanValueString(bool bv)
        {
            return bv ? "TRUE_CUSTOM" : "FALSE_CUSTOM";
        }

        public override string GetErrorValueString(string err)
        {
            // Simple mapping for demonstration.
            return err.Replace("#", "#CUST_");
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the sample workbook.
            const string filePath = "sample.xlsx";

            // -----------------------------------------------------------------
            // 1. Create a sample workbook (if it does not already exist) and save it.
            // -----------------------------------------------------------------
            if (!System.IO.File.Exists(filePath))
            {
                // Create a new workbook – using the provided constructor rule.
                Workbook wbCreate = new Workbook();

                // Populate some data.
                Worksheet ws = wbCreate.Worksheets[0];
                Cells cells = ws.Cells;
                for (int row = 0; row < 1000; row++)
                {
                    cells[row, 0].PutValue($"Item {row}");
                    cells[row, 1].PutValue(row * 1.1);
                }

                // Save the workbook – using the provided Save method.
                wbCreate.Save(filePath);
                Console.WriteLine($"Sample workbook created at '{filePath}'.");
            }

            // -----------------------------------------------------------------
            // 2. Load without custom GlobalizationSettings and measure time.
            // -----------------------------------------------------------------
            Stopwatch sw = new Stopwatch();
            sw.Start();

            // Load workbook using the constructor that accepts a file path.
            Workbook wbDefault = new Workbook(filePath);

            sw.Stop();
            long loadTimeDefaultMs = sw.ElapsedMilliseconds;
            Console.WriteLine($"Load time without custom globalization: {loadTimeDefaultMs} ms");

            // -----------------------------------------------------------------
            // 3. Load and then apply custom GlobalizationSettings; measure both phases.
            // -----------------------------------------------------------------
            // Measure load time (same as above, but separate stopwatch for clarity).
            sw.Restart();
            Workbook wbWithCustom = new Workbook(filePath);
            sw.Stop();
            long loadTimeWithCustomMs = sw.ElapsedMilliseconds;
            Console.WriteLine($"Load time before applying custom globalization: {loadTimeWithCustomMs} ms");

            // Measure the time required to assign the custom settings.
            sw.Restart();
            wbWithCustom.Settings.GlobalizationSettings = new CustomGlobalizationSettings();
            sw.Stop();
            long applySettingsTimeMs = sw.ElapsedMilliseconds;
            Console.WriteLine($"Time to apply custom globalization settings: {applySettingsTimeMs} ms");

            // -----------------------------------------------------------------
            // 4. Optional: demonstrate that the custom settings affect cell display.
            // -----------------------------------------------------------------
            // Put a boolean value and an error value to see custom strings.
            Cells demoCells = wbWithCustom.Worksheets[0].Cells;
            demoCells[0, 2].PutValue(true);
            demoCells[1, 2].PutValue("#DIV/0!");

            // Access the string representations (they will use the custom settings).
            string boolStr = demoCells[0, 2].StringValue; // Expected "TRUE_CUSTOM"
            string errStr = demoCells[1, 2].StringValue; // Expected "#CUST_DIV/0!"

            Console.WriteLine($"Custom boolean string: {boolStr}");
            Console.WriteLine($"Custom error string: {errStr}");

            // -----------------------------------------------------------------
            // 5. Summary output.
            // -----------------------------------------------------------------
            Console.WriteLine();
            Console.WriteLine("Performance Summary:");
            Console.WriteLine($"- Default load time          : {loadTimeDefaultMs} ms");
            Console.WriteLine($"- Load time (custom later)   : {loadTimeWithCustomMs} ms");
            Console.WriteLine($"- Apply custom settings time : {applySettingsTimeMs} ms");
        }
    }
}