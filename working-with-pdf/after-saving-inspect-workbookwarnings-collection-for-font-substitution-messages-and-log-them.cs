using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Put some text into a cell and assign a font that likely does not exist
                Cell cell = sheet.Cells["A1"];
                cell.PutValue("Text with a missing font");
                Style style = workbook.CreateStyle();
                style.Font.Name = "NonExistentFont";
                cell.SetStyle(style);

                // Define output file path
                string outputPath = "Output.xlsx";

                // Ensure the directory for the output file exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);

                // NOTE: Warning retrieval APIs may not be available in all Aspose.Cells versions.
                // If needed, use workbook.GetWarnings() in a version that supports it.
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}