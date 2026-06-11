using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: use provided constructor)
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // Prepare sample worksheets to demonstrate classification
            // -------------------------------------------------

            // Worksheet 0: contains only cell data
            Worksheet dataOnlySheet = workbook.Worksheets[0];
            dataOnlySheet.Name = "DataOnly";
            dataOnlySheet.Cells["A1"].PutValue("Sample text");
            dataOnlySheet.Cells["B2"].PutValue(12345);

            // Worksheet 1: contains only shapes
            Worksheet shapeOnlySheet = workbook.Worksheets.Add("ShapeOnly");
            // Add a rectangle shape
            shapeOnlySheet.Shapes.AddRectangle(1, 1, 0, 0, 120, 60);

            // Worksheet 2: contains both data and shapes
            Worksheet mixedSheet = workbook.Worksheets.Add("Mixed");
            mixedSheet.Cells["C3"].PutValue(DateTime.Now);
            // Add an ellipse shape using AddShape with MsoDrawingType.Oval (ellipse)
            mixedSheet.Shapes.AddShape(MsoDrawingType.Oval, 1, 1, 0, 0, 80, 80);

            // -------------------------------------------------
            // Classification logic: data‑only, shape‑only, mixed, or empty
            // -------------------------------------------------
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine if the worksheet has any cell data
                bool hasData = sheet.Cells.MaxDataRow >= 0 && sheet.Cells.MaxDataColumn >= 0;

                // Determine if the worksheet has any drawing shapes
                bool hasShapes = sheet.Shapes.Count > 0;

                string classification = hasData && hasShapes ? "Mixed content"
                                    : hasData ? "Data‑only"
                                    : hasShapes ? "Shape‑only"
                                    : "Empty";

                Console.WriteLine($"Worksheet \"{sheet.Name}\": {classification}");
            }

            // Save the workbook (lifecycle rule: use provided Save method)
            string outputPath = "ClassifiedWorkbook.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{Path.GetFullPath(outputPath)}\"");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}