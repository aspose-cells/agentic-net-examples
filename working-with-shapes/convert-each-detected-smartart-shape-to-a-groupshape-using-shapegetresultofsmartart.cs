using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;
using Aspose.Cells.Drawing;

namespace AsposeCellsSmartArtConversion
{
    public class SmartArtToGroupShapeConverter
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook that contains SmartArt shapes
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Get the collection of shapes on the current worksheet
                    ShapeCollection shapes = sheet.Shapes;

                    // Loop through each shape in the collection
                    for (int i = 0; i < shapes.Count; i++)
                    {
                        Shape shape = shapes[i];

                        // Check if the shape is a SmartArt object
                        if (shape.IsSmartArt)
                        {
                            // Convert the SmartArt shape to a GroupShape
                            GroupShape groupShape = shape.GetResultOfSmartArt();

                            // Optional manipulation of the resulting group shape
                            if (groupShape != null)
                            {
                                groupShape.Left = shape.Left + 20;
                                groupShape.Top = shape.Top + 20;
                            }
                        }
                    }
                }

                // Enable updating of SmartArt when saving
                OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
                {
                    UpdateSmartArt = true
                };

                // Save the modified workbook
                workbook.Save(outputPath, saveOptions);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Runtime error: {ex.Message}");
            }
        }
    }
}