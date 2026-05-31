using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class DisableShapeReflections
    {
        /// <summary>
        /// Disables reflection effects on all shapes in a workbook.
        /// If "input.xlsx" exists in the executable folder it is loaded; otherwise a new workbook is created.
        /// </summary>
        public static void Run()
        {
            try
            {
                const string inputPath = "input.xlsx";
                Workbook workbook;

                // Load existing workbook if the file is present; otherwise create a new one.
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                }

                // Iterate through all worksheets and their shapes.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (Shape shape in sheet.Shapes)
                    {
                        // Disable reflection by setting its type to None.
                        shape.Reflection.Type = ReflectionEffectType.None;
                    }
                }

                // Save the modified workbook.
                const string outputPath = "Workbook_NoReflections.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point required for console applications.
    public class Program
    {
        public static void Main(string[] args)
        {
            DisableShapeReflections.Run();
        }
    }
}