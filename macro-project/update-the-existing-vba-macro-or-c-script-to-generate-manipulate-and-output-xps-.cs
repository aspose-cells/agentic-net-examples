using System;
using Aspose.Cells;

namespace AsposeCellsXpsDemo
{
    public class XpsGeneration
    {
        public static void Run()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Aspose.Cells XPS Demo");
            sheet.Cells["A2"].PutValue(DateTime.Now);
            sheet.Cells["B1"].PutValue(12345);
            sheet.Cells["B2"].PutValue(67.89);

            // OPTIONAL: Remove any VBA/macros if present
            workbook.RemoveMacro();

            // Configure XPS save options
            XpsSaveOptions saveOptions = new XpsSaveOptions
            {
                OnePagePerSheet = true,
                DefaultFont = "Arial",
                CheckFontCompatibility = true,
                CheckWorkbookDefaultFont = true,
                AllColumnsInOnePagePerSheet = true,
                PageIndex = 0,
                PageCount = 1
            };

            // Save the workbook as XPS
            string outputPath = "GeneratedReport.xps";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook successfully saved as XPS to '{outputPath}'.");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            XpsGeneration.Run();
        }
    }
}