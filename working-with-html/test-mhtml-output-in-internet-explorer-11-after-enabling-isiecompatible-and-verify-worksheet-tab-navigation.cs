using System;
using Aspose.Cells;

namespace AsposeCellsMhtmlIeTest
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook with two worksheets
            Workbook workbook = new Workbook();
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "FirstSheet";
            sheet1.Cells["A1"].PutValue("Sheet 1 - Cell A1");
            sheet1.Cells["B2"].PutValue("Sheet 1 - Cell B2");

            Worksheet sheet2 = workbook.Worksheets.Add("SecondSheet");
            sheet2.Cells["A1"].PutValue("Sheet 2 - Cell A1");
            sheet2.Cells["B2"].PutValue("Sheet 2 - Cell B2");

            // Ensure workbook tabs are visible (required for tab navigation in HTML/MHTML)
            workbook.Settings.ShowTabs = true;

            // Configure HTML save options for MHTML output
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                // Enable IE compatibility mode so the generated MHTML works in Internet Explorer 11
                IsIECompatible = true,

                // Export all worksheets (default) and keep the tab strip
                ExportSingleTab = false,

                // Export worksheet properties (including tab information)
                ExportWorksheetProperties = true
            };

            // Save the workbook as MHTML (MHT) using the configured options
            string outputPath = "Workbook_IECompatible.mht";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"MHTML file saved to '{outputPath}'. Open it in Internet Explorer 11 to verify tab navigation.");
        }
    }
}