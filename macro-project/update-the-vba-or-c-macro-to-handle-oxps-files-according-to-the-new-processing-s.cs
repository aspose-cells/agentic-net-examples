using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    public class OXpsProcessingDemo
    {
        public static void Run()
        {
            // Load a macro‑enabled workbook (e.g., .xlsm)
            Workbook workbook = new Workbook("MacroEnabledWorkbook.xlsm");

            // Remove any VBA/macros from the workbook
            workbook.RemoveMacro();

            // Configure XPS save options for OpenXPS (OXPS) output
            XpsSaveOptions saveOptions = new XpsSaveOptions
            {
                // Ensure each sheet is rendered on a single page
                OnePagePerSheet = true,

                // Use a default font that supports Unicode characters
                DefaultFont = "Arial",

                // Render all columns of a sheet on one page (optional)
                AllColumnsInOnePagePerSheet = true,

                // Set page range (first page only in this example)
                PageIndex = 0,
                PageCount = 1,

                // Enable font compatibility checks
                CheckFontCompatibility = true,
                CheckWorkbookDefaultFont = true,

                // Keep gridlines visible
                GridlineType = GridlineType.Dotted
            };

            // Save the workbook as an OpenXPS file with the specified options
            workbook.Save("ProcessedDocument.oxps", saveOptions);

            Console.WriteLine("Workbook processed and saved as OXPS successfully.");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            OXpsProcessingDemo.Run();
        }
    }
}