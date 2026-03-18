using System;
using Aspose.Cells;

namespace AsposeCellsToPptx
{
    public class Converter
    {
        public static void Run()
        {
            // Path to the source Excel workbook
            string sourcePath = "input.xlsx";

            // Desired path for the generated PowerPoint presentation
            string destinationPath = "output.pptx";

            // Load the Excel workbook from the file system
            Workbook workbook = new Workbook(sourcePath);

            // Create PPTX save options (default settings)
            PptxSaveOptions saveOptions = new PptxSaveOptions();

            // Save the workbook as a PPTX file using the specified options
            workbook.Save(destinationPath, saveOptions);

            Console.WriteLine($"Conversion completed: '{sourcePath}' -> '{destinationPath}'");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            Converter.Run();
        }
    }
}