// Title: Create an Excel workbook in C# with Aspose.Cells that automatically inserts page breaks based on row height
// AI Prompts: Write C# code using Aspose.Cells to build a worksheet, fill it with data, and configure the PageSetup so Excel adds page breaks automatically when printed. | Show how to set PageSetup.FitToPagesWide = 1 and FitToPagesTall = 0 in Aspose.Cells to achieve dynamic pagination for a generated XLSX file.
// Common Searches: Aspose.Cells C# set page setup for automatic pagination | How to enable dynamic page breaks in Excel using Aspose.Cells .NET | FitToPagesWide 1 FitToPagesTall 0 example Aspose.Cells | Create Excel file with auto page breaks based on content height C# | Aspose.Cells print layout page breaks configuration
// Tags: Aspose.Cells page setup auto pagination | C# FitToPagesWide property Aspose.Cells | FitToPagesTall zero dynamic breaks | Generate XLSX with Aspose.Cells | Excel print layout configuration Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPageBreakDemo
{
    // The sample creates a new workbook, renames the first worksheet, populates 200 rows and 10 columns with sample data, configures PageSetup to fit the content to one page width while leaving height unrestricted (FitToPagesWide = 1, FitToPagesTall = 0) so Excel inserts page breaks automatically, and saves the result as AutomaticPageBreaks.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (or load an existing one if needed)
                Workbook workbook = new Workbook(); // default workbook with one worksheet

                // Access the first worksheet and rename it
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Data";

                // Populate the worksheet with sample data
                for (int row = 0; row < 200; row++)
                {
                    for (int col = 0; col < 10; col++)
                    {
                        sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                    }
                }

                // Configure page setup for automatic page breaks
                PageSetup pageSetup = sheet.PageSetup;
                // Automatic page breaks are enabled by default; no explicit property needed
                pageSetup.FitToPagesTall = 0;   // 0 disables fitting to a specific number of pages tall
                pageSetup.FitToPagesWide = 1;   // fit content to one page width

                // Save the workbook
                string outputPath = "AutomaticPageBreaks.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
