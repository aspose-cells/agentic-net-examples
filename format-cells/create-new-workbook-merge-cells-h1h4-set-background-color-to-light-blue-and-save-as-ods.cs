using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Ods;

namespace AsposeCellsExamples
{
    public class MergeCellsAndSaveOds
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook with a default worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Merge cells H1:H4 (zero‑based indices: row 0, column 7, 4 rows, 1 column)
                worksheet.Cells.Merge(0, 7, 4, 1);

                // Apply light‑blue background to the merged cell
                Style style = worksheet.Cells["H1"].GetStyle();
                style.ForegroundColor = Color.LightBlue;
                style.Pattern = BackgroundType.Solid;
                worksheet.Cells["H1"].SetStyle(style);

                // Prepare ODS save options (optional: set generator type)
                OdsSaveOptions saveOptions = new OdsSaveOptions
                {
                    GeneratorType = OdsGeneratorType.LibreOffice
                };

                // Define output file path
                string outputPath = "MergedCellsLightBlue.ods";

                // Save the workbook as ODS
                workbook.Save(outputPath, saveOptions);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            MergeCellsAndSaveOds.Run();
        }
    }
}