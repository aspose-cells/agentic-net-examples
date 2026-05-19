using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeAdjustLogging
{
    class Program
    {
        static void Main()
        {
            const string logPath = "ShapeAdjustmentsAudit.log";

            try
            {
                // Clear previous log content
                File.WriteAllText(logPath, string.Empty);

                // Create a new workbook and obtain the first worksheet
                var workbook = new Workbook();
                var worksheet = workbook.Worksheets[0];

                // Add an auto shape that supports adjustment guides
                var shape = worksheet.Shapes.AddAutoShape(
                    AutoShapeType.RightArrowCallout, 2, 0, 2, 0, 200, 150);

                // Access the shape's geometry (contains adjustment guides)
                var geometry = shape.Geometry;

                // Ensure there are some adjustment guides to work with
                if (geometry.ShapeAdjustValues.Count == 0)
                {
                    geometry.ShapeAdjustValues.Add("adj1", 10.0);
                    geometry.ShapeAdjustValues.Add("adj2", 20.0);
                }

                // Log each adjustment change
                using (var logWriter = new StreamWriter(logPath, true))
                {
                    for (int i = 0; i < geometry.ShapeAdjustValues.Count; i++)
                    {
                        double originalValue = geometry.ShapeAdjustValues[i].Value;
                        double newValue = originalValue + 5.0;

                        // Apply the new adjustment value
                        geometry.ShapeAdjustValues[i].Value = newValue;

                        // ShapeGuide does not expose a Name property; use a fallback name
                        string guideName = $"adj{i + 1}";

                        logWriter.WriteLine(
                            $"{DateTime.Now:u} - Shape '{shape.Name}' guide '{guideName}' changed from {originalValue} to {newValue}");
                    }
                }

                // Save the workbook
                const string outputFile = "ShapeAdjustmentsDemo.xlsx";
                try
                {
                    workbook.Save(outputFile);
                }
                catch (Exception saveEx)
                {
                    Console.Error.WriteLine($"Error saving workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                // Log unexpected errors
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}