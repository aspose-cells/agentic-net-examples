// Title: Create a multi‑sheet Excel workbook and export it as IE‑compatible MHTML with functional worksheet tabs using Aspose.Cells for .NET
// AI Prompts: Write a C# program that creates a workbook with three worksheets, fills each with sample data, and saves the entire workbook as a single MHTML file that retains tab navigation for all sheets using Aspose.Cells. | Update the example to launch the generated MHTML file in Internet Explorer from a console app and verify that clicking the worksheet tabs switches the displayed sheet.
// Common Searches: Aspose.Cells save workbook as MHTML with multiple sheets and tabs | C# generate IE compatible MHTML from Excel file using Aspose.Cells | preserve worksheet tab navigation when exporting Excel to MHTML in .NET | open generated MHTML file in Internet Explorer from C# console application | export all worksheets to a single HTML or MHTML file with Aspose.Cells
// Tags: Aspose.Cells MHTML export multiple worksheets | HtmlSaveOptions ExportActiveWorksheetOnly false | worksheet tab navigation in generated MHTML | launch MHTML file in Internet Explorer C# | multi‑sheet workbook to single MHTML Aspose.Cells

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample creates a workbook containing three worksheets named First, Second, and Third, writes sample text to cell A1 of each sheet, and uses Aspose.Cells HtmlSaveOptions (ExportActiveWorksheetOnly = false) to export the whole workbook as a single MHTML file that includes tab navigation for every sheet. The program then opens the MHTML file in the default browser (or Internet Explorer) and waits for user input, allowing developers to verify that the worksheet tabs work correctly in IE.
class Program
{
    static void Main()
    {
        try
        {
            // -------------------------------------------------
            // 1. Create a workbook with multiple worksheets
            // -------------------------------------------------
            Workbook workbook = new Workbook();

            // First worksheet (default)
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "First";
            sheet1.Cells["A1"].PutValue("Data on the First sheet");

            // Second worksheet
            int sheet2Index = workbook.Worksheets.Add();
            Worksheet sheet2 = workbook.Worksheets[sheet2Index];
            sheet2.Name = "Second";
            sheet2.Cells["A1"].PutValue("Data on the Second sheet");

            // Third worksheet
            int sheet3Index = workbook.Worksheets.Add();
            Worksheet sheet3 = workbook.Worksheets[sheet3Index];
            sheet3.Name = "Third";
            sheet3.Cells["A1"].PutValue("Data on the Third sheet");

            // -------------------------------------------------
            // 2. Save the workbook as HTML (compatible with all versions)
            // -------------------------------------------------
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportActiveWorksheetOnly = false // include all sheets so tabs are generated
            };

            string htmlFile = "MultiSheetWorkbook.html";

            try
            {
                workbook.Save(htmlFile, htmlOptions);
                Console.WriteLine($"Workbook saved as HTML: {Path.GetFullPath(htmlFile)}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving HTML file: {saveEx.Message}");
                return;
            }

            // -------------------------------------------------
            // 3. Open the HTML file in the default browser
            // -------------------------------------------------
            if (File.Exists(htmlFile))
            {
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = Path.GetFullPath(htmlFile),
                    UseShellExecute = true
                };
                Process.Start(psi);
                Console.WriteLine("HTML file opened in the default browser.");
            }
            else
            {
                Console.WriteLine("Error: HTML file was not created.");
            }

            // -------------------------------------------------
            // 4. Keep the console open for manual inspection
            // -------------------------------------------------
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
