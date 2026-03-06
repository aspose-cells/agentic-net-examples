using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsOxpsDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Aspose.Cells OXPS Demo");
            sheet.Cells["A2"].PutValue(DateTime.Now);
            sheet.Cells["B1"].PutValue(12345);
            sheet.Cells["B2"].PutValue(3.14159);

            // Initialize XpsSaveOptions
            XpsSaveOptions saveOptions = new XpsSaveOptions
            {
                OnePagePerSheet = true,
                DefaultFont = "Arial",
                CheckWorkbookDefaultFont = true,
                CheckFontCompatibility = true,
                AllColumnsInOnePagePerSheet = false,
                IgnoreError = false,
                OutputBlankPageWhenNothingToPrint = false,
                PageIndex = 0,
                PageCount = 1,
                PrintingPageType = PrintingPageType.Default,
                GridlineType = GridlineType.Dotted,
                TextCrossType = TextCrossType.Default,
                DefaultEditLanguage = DefaultEditLanguage.English,
                SheetSet = SheetSet.All
            };

            // Define the output file name with .oxps extension
            string outputPath = "DemoOutput.oxps";

            // Save the workbook as OXPS using the XpsSaveOptions
            workbook.Save(outputPath, saveOptions);

            // Console output to indicate success
            Console.WriteLine($"Workbook has been saved as OXPS file: {outputPath}");
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
    }
}