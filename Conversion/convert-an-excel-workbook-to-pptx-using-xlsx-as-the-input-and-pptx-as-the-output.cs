using System;
using Aspose.Cells;               // Core Aspose.Cells namespace
using Aspose.Cells.Slides;       // Required for SlideViewType enum (optional)

namespace AsposeCellsConversionDemo
{
    public class ExcelToPptxConverter
    {
        public static void Run()
        {
            // Path to the source Excel file (XLSX)
            string sourcePath = "input.xlsx";

            // Path for the resulting PowerPoint file (PPTX)
            string destPath = "output.pptx";

            // Load the Excel workbook from the file system
            Workbook workbook = new Workbook(sourcePath);

            // Create PPTX save options (default constructor)
            PptxSaveOptions saveOptions = new PptxSaveOptions();

            // Optional: customize options, e.g., ignore hidden rows and export as view
            saveOptions.IgnoreHiddenRows = true;
            saveOptions.ExportViewType = SlideViewType.View;

            // Save the workbook as a PPTX file using the specified options
            workbook.Save(destPath, saveOptions);

            Console.WriteLine($"Conversion completed: '{sourcePath}' -> '{destPath}'");
        }
    }

    // Entry point for demonstration
    class Program
    {
        static void Main()
        {
            ExcelToPptxConverter.Run();
        }
    }
}