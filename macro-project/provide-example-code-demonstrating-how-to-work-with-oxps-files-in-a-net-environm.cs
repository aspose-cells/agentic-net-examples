using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsOxpsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "DemoSheet";
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["A2"].PutValue("World");
            sheet.Cells["B1"].PutValue(12345);

            // Configure XpsSaveOptions for OXPS output
            XpsSaveOptions saveOptions = new XpsSaveOptions();
            saveOptions.OnePagePerSheet = true;          // each sheet on a single page
            saveOptions.DefaultFont = "Arial";           // default font for Unicode characters
            saveOptions.CheckWorkbookDefaultFont = true; // use workbook default font when needed
            saveOptions.CheckFontCompatibility = true;  // verify font compatibility

            // Save the workbook directly as an OXPS file
            string oxpsFile = "DemoOutput.oxps";
            workbook.Save(oxpsFile, saveOptions);

            // -----------------------------------------------------------------
            // Demonstrate conversion from an existing XLSX file to OXPS using
            // ConversionUtility with explicit load and save options.
            // -----------------------------------------------------------------

            // Prepare a source XLSX file (reuse the workbook created above)
            string sourceXlsx = "SourceWorkbook.xlsx";
            workbook.Save(sourceXlsx, SaveFormat.Xlsx);

            // Destination OXPS file path
            string destOxps = "ConvertedOutput.oxps";

            // Load options for the source XLSX (default settings)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // Save options for the OXPS conversion
            XpsSaveOptions convertSaveOptions = new XpsSaveOptions();
            convertSaveOptions.OnePagePerSheet = false; // allow multiple pages per sheet
            convertSaveOptions.DefaultFont = "Times New Roman";

            // Perform the conversion
            ConversionUtility.Convert(sourceXlsx, loadOptions, destOxps, convertSaveOptions);

            Console.WriteLine("OXPS files have been created and conversion completed successfully.");
        }
    }
}