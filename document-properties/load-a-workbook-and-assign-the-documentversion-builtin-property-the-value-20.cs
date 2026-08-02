using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class SetDocumentVersion
    {
        public static void Run()
        {
            try
            {
                string inputPath = "input.xlsx";

                // Ensure the input file exists; create a blank workbook if it does not.
                if (!File.Exists(inputPath))
                {
                    Workbook blank = new Workbook();
                    blank.Save(inputPath, SaveFormat.Xlsx);
                }

                // Load the workbook.
                Workbook workbook = new Workbook(inputPath);

                // Set the built‑in DocumentVersion property.
                workbook.BuiltInDocumentProperties.DocumentVersion = "2.0";

                // Save the modified workbook.
                string outputPath = "output.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            SetDocumentVersion.Run();
        }
    }
}