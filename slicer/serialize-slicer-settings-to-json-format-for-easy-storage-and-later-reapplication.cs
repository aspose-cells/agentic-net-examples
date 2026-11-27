// Title: Serialize Aspose.Cells slicer properties to JSON and reapply them in a .NET workbook
// Common Searches: C# Aspose.Cells save slicer configuration to a JSON file | load slicer settings from JSON using Aspose.Cells .NET | Aspose.Cells serialize slicer style and layout to JSON | how to persist slicer position and style in Aspose.Cells workbook | restore pivot table slicer properties from JSON in C# | Aspose.Cells slicer settings POCO serialization example
// Tags: Aspose.Cells slicer JSON serialization | C# capture slicer properties Aspose.Cells | deserialize slicer settings Aspose.Cells | persist slicer layout Aspose.Cells workbook | apply slicer style from JSON Aspose.Cells | pivot table slicer configuration storage .NET

using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace SlicerSettingsSerialization
{
    // Simple POCO that mirrors the slicer properties we want to persist
    // The example creates a workbook with sample data, adds a pivot table and a slicer, captures the slicer's visual and style attributes into a POCO, serializes this object to an indented JSON file, reads the JSON back, applies the deserialized settings to a second slicer, and saves the workbook. It demonstrates how to persist and restore Aspose.Cells slicer settings using C# and System.Text.Json.
    public class SlicerSettings
    {
        public string Caption { get; set; }
        public int TopPixel { get; set; }
        public int LeftPixel { get; set; }
        public int HeightPixel { get; set; }
        public int WidthPixel { get; set; }
        public int NumberOfColumns { get; set; }
        public string StyleType { get; set; } // store enum name for easy deserialization
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Paths for temporary files
            string jsonPath = "slicerSettings.json";
            string outputPath = "WorkbookWithRestoredSlicer.xlsx";

            try
            {
                // ------------------------------------------------------------
                // 1. Create a workbook, add data, a pivot table and a slicer
                // ------------------------------------------------------------
                Workbook wb = new Workbook();
                Worksheet ws = wb.Worksheets[0];
                Cells cells = ws.Cells;

                // Sample data
                cells["A1"].PutValue("Category");
                cells["B1"].PutValue("Product");
                cells["C1"].PutValue("Quantity");

                cells["A2"].PutValue("Fruits");
                cells["B2"].PutValue("Apple");
                cells["C2"].PutValue(10);

                cells["A3"].PutValue("Fruits");
                cells["B3"].PutValue("Banana");
                cells["C3"].PutValue(15);

                cells["A4"].PutValue("Vegetables");
                cells["B4"].PutValue("Carrot");
                cells["C4"].PutValue(8);

                // Create a pivot table based on the data range
                int pivotIdx = ws.PivotTables.Add("A1:C4", "E1", "PivotTable1");
                PivotTable pivot = ws.PivotTables[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Column, "Product");
                pivot.AddFieldToArea(PivotFieldType.Data, "Quantity");
                pivot.RefreshData();
                pivot.CalculateData();

                // Add a slicer linked to the "Category" field of the pivot table
                // Destination cell must be a valid address (e.g., "G1")
                int slicerIdx = ws.Slicers.Add(pivot, "G1", "Category");
                Slicer slicer = ws.Slicers[slicerIdx];

                // Configure slicer appearance
                slicer.Caption = "Select Category";
                slicer.TopPixel = 100;
                slicer.LeftPixel = 400;
                slicer.HeightPixel = 150;
                slicer.WidthPixel = 200;
                slicer.NumberOfColumns = 2;
                slicer.StyleType = SlicerStyleType.SlicerStyleDark2;

                // ------------------------------------------------------------
                // 2. Capture slicer settings into a POCO and serialize to JSON
                // ------------------------------------------------------------
                SlicerSettings settings = CaptureSlicerSettings(slicer);
                string json = JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(jsonPath, json);
                Console.WriteLine($"Slicer settings serialized to '{jsonPath}':");
                Console.WriteLine(json);

                // ------------------------------------------------------------
                // 3. Demonstrate re‑application: create a second slicer and apply saved settings
                // ------------------------------------------------------------
                // Add a second slicer (initially with default values)
                int slicerIdx2 = ws.Slicers.Add(pivot, "G10", "Category");
                Slicer restoredSlicer = ws.Slicers[slicerIdx2];

                // Load settings from JSON (ensure the file exists)
                if (!File.Exists(jsonPath))
                    throw new FileNotFoundException($"Settings file not found: {jsonPath}");

                string loadedJson = File.ReadAllText(jsonPath);
                SlicerSettings loadedSettings = JsonSerializer.Deserialize<SlicerSettings>(loadedJson);

                // Apply the loaded settings to the new slicer
                ApplySlicerSettings(restoredSlicer, loadedSettings);

                // ------------------------------------------------------------
                // 4. Save the workbook to verify that the restored slicer matches the original
                // ------------------------------------------------------------
                wb.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved as '{outputPath}'. The restored slicer should mirror the original settings.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Extracts the relevant properties from a Slicer into a plain object
        private static SlicerSettings CaptureSlicerSettings(Slicer slicer)
        {
            return new SlicerSettings
            {
                Caption = slicer.Caption,
                TopPixel = slicer.TopPixel,
                LeftPixel = slicer.LeftPixel,
                HeightPixel = slicer.HeightPixel,
                WidthPixel = slicer.WidthPixel,
                NumberOfColumns = slicer.NumberOfColumns,
                StyleType = slicer.StyleType.ToString()
            };
        }

        // Applies a previously captured settings object back onto a Slicer instance
        private static void ApplySlicerSettings(Slicer slicer, SlicerSettings settings)
        {
            if (settings == null) throw new ArgumentNullException(nameof(settings));

            slicer.Caption = settings.Caption;
            slicer.TopPixel = settings.TopPixel;
            slicer.LeftPixel = settings.LeftPixel;
            slicer.HeightPixel = settings.HeightPixel;
            slicer.WidthPixel = settings.WidthPixel;
            slicer.NumberOfColumns = settings.NumberOfColumns;

            // Parse the enum name back to SlicerStyleType
            if (Enum.TryParse<SlicerStyleType>(settings.StyleType, out var style))
            {
                slicer.StyleType = style;
            }
            else
            {
                Console.WriteLine($"Warning: Unable to parse StyleType '{settings.StyleType}'. Keeping default style.");
            }
        }
    }
}
