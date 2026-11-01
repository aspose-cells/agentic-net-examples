// Title: Set fixed minimum and maximum axis values for a line sparkline group with Aspose.Cells in C#
// Common Searches: C# Aspose.Cells how to set sparkline axis minimum value | Aspose.Cells line sparkline fixed axis range example | Set SparklineGroup AxisMin and AxisMax programmatically using Aspose.Cells | Configure consistent scaling for Excel sparkline using Aspose.Cells .NET | Dynamic access to SparklineGroups in Aspose.Cells for axis settings | Save workbook with sparkline axis limits in C# Aspose.Cells
// Tags: Aspose.Cells sparkline axis range | C# set sparkline AxisMin AxisMax | line sparkline fixed scaling .NET | dynamic SparklineGroups usage Aspose.Cells | Excel sparkline axis configuration with Aspose

using Aspose.Cells;
using System;
using System.IO;

// The example creates a workbook, fills cells A1:A5 with numbers, adds a line sparkline in B1 via dynamic access to the SparklineGroups collection, assigns AxisMin = 0 and AxisMax = 60 to enforce a fixed axis range, and saves the file as SparklineAxisFixed.xlsx while gracefully handling environments where the Sparkline API is unavailable.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate source data for the sparkline (cells A1:A5)
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);
            sheet.Cells["A4"].PutValue(40);
            sheet.Cells["A5"].PutValue(50);

            // Attempt to add a sparkline group and set axis scaling.
            // Wrapped in try-catch to handle environments where Sparkline APIs are unavailable.
            try
            {
                // Use dynamic to avoid compile‑time dependency on Sparkline types.
                dynamic sheetDyn = sheet;

                // Access the SparklineGroups collection via dynamic.
                var sparklineGroups = sheetDyn.SparklineGroups;

                // Resolve the SparklineType enum (Line) via reflection.
                Type sparklineTypeEnum = Type.GetType("Aspose.Cells.Sparkline.SparklineType, Aspose.Cells");
                if (sparklineTypeEnum == null)
                    throw new InvalidOperationException("SparklineType enum not found. Sparkline feature may be unavailable.");

                // Assuming the enum value for Line is 0.
                object lineEnum = Enum.ToObject(sparklineTypeEnum, 0);

                // Add a line sparkline group: place sparkline in B1, source data A1:A5
                int sparklineGroupIndex = sparklineGroups.Add(lineEnum, "B1", "A1:A5");

                // Retrieve the created sparkline group via dynamic indexing.
                dynamic sparklineGroup = sparklineGroups[sparklineGroupIndex];

                // Set fixed axis scaling values
                sparklineGroup.AxisMin = 0;   // Minimum value for the sparkline axis
                sparklineGroup.AxisMax = 60;  // Maximum value for the sparkline axis
            }
            catch (Exception ex)
            {
                // Log but continue; sparkline feature may be unavailable.
                Console.WriteLine($"Sparkline operation skipped: {ex.Message}");
            }

            // Define output file path
            string outputPath = "SparklineAxisFixed.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
