// Title: How to add multiple data ranges to a single SparklineGroup for a multi‑series line sparkline using Aspose.Cells in C#
// AI Prompts: Create a line‑type SparklineGroup at D1:D5, assign the first data range A1:A5, then add additional data ranges B1:B5 and C1:C5, enable markers and set line weight, using Aspose.Cells for .NET. | Invoke SparklineGroup.Add and SparklineGroup.AddDataRange via reflection so the code compiles even when the Sparkline assembly is not referenced, and save the workbook as an .xlsx file.
// Common Searches: asp.net add multiple data ranges to a sparkline group Aspose.Cells example | c# create multi series line sparkline with Aspose.Cells | using reflection to add sparkline in Aspose.Cells when Sparkline assembly missing | set ShowMarkers and LineWeight for sparkline in Aspose.Cells .NET | generate multi‑series sparkline in Excel using Aspose.Cells C#
// Tags: Aspose.Cells add data range to SparklineGroup | multi-series line sparkline C# | SparklineGroup reflection invocation Aspose | configure ShowMarkers LineWeight Aspose.Cells | save workbook as .xlsx Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, fills columns A‑C with sample numeric data, then uses reflection to add a line‑type SparklineGroup at D1:D5 with the initial range A1:A5. It adds two more ranges (B1:B5 and C1:C5) to produce a multi‑series sparkline, enables markers, adjusts line weight, and saves the file as MultiSeriesSparkline.xlsx, while gracefully handling environments where the Sparkline feature is unavailable.
class SparklineMultiSeriesExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for three series (columns A, B, C)
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);
            sheet.Cells["A4"].PutValue(25);
            sheet.Cells["A5"].PutValue(15);

            sheet.Cells["B1"].PutValue(12);
            sheet.Cells["B2"].PutValue(22);
            sheet.Cells["B3"].PutValue(28);
            sheet.Cells["B4"].PutValue(27);
            sheet.Cells["B5"].PutValue(18);

            sheet.Cells["C1"].PutValue(8);
            sheet.Cells["C2"].PutValue(18);
            sheet.Cells["C3"].PutValue(35);
            sheet.Cells["C4"].PutValue(20);
            sheet.Cells["C5"].PutValue(10);

            // Attempt to create a multi‑series sparkline using reflection.
            // This avoids compile‑time dependency on the Sparkline namespace,
            // allowing the code to compile even if the Sparkline assembly is absent.
            try
            {
                // Resolve SparklineType enum (Line)
                Type sparklineTypeEnum = Type.GetType("Aspose.Cells.Sparkline.SparklineType, Aspose.Cells");
                object lineEnumValue = Enum.Parse(sparklineTypeEnum, "Line");

                // Add a sparkline group (location D1:D5, first data range A1:A5)
                var sparklineGroups = sheet.SparklineGroups;
                var addMethod = sparklineGroups.GetType().GetMethod("Add");
                int groupIndex = (int)addMethod.Invoke(sparklineGroups, new object[] { lineEnumValue, "D1:D5", "A1:A5" });

                // Retrieve the created SparklineGroup
                var sparklineGroup = sparklineGroups[groupIndex];

                // Add additional data ranges (B1:B5 and C1:C5)
                var addDataRangeMethod = sparklineGroup.GetType().GetMethod("AddDataRange");
                addDataRangeMethod.Invoke(sparklineGroup, new object[] { "B1:B5" });
                addDataRangeMethod.Invoke(sparklineGroup, new object[] { "C1:C5" });

                // Set optional style properties
                var showMarkersProp = sparklineGroup.GetType().GetProperty("ShowMarkers");
                showMarkersProp.SetValue(sparklineGroup, true);

                var lineWeightProp = sparklineGroup.GetType().GetProperty("LineWeight");
                lineWeightProp.SetValue(sparklineGroup, 0.75);
            }
            catch (Exception ex)
            {
                // Sparkline feature may not be available in the referenced Aspose.Cells version
                Console.WriteLine("Sparkline creation skipped: " + ex.Message);
            }

            // Define output file path
            string outputPath = "MultiSeriesSparkline.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
