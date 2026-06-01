using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsJapaneseEraPdf
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Set the regional settings to Japan so that Japanese era formatting can be applied
            // (uses WorkbookSettings.Region property)
            workbook.Settings.Region = CountryCode.Japan;

            // Access the first worksheet and a target cell
            Worksheet sheet = workbook.Worksheets[0];
            Cell cell = sheet.Cells[0, 0]; // A1

            // Put an Excel serial date value (e.g., 44089 corresponds to 2020-09-15)
            cell.PutValue(44089);

            // Retrieve the cell's style, set a custom format that displays the date in Japanese era
            // The format string uses the locale identifier for Japanese ([$-F800])
            Style style = cell.GetStyle();
            style.Custom = "[$-F800]ggge\"年\"m\"月\"d\"日\""; // Example: Reiwa3年9月15日
            cell.SetStyle(style);

            // Configure PDF save options, specifying a Japanese font to ensure proper rendering
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                DefaultFont = "MS Gothic", // Japanese font
                CheckWorkbookDefaultFont = true
            };

            // Save the workbook as PDF (lifecycle: save)
            workbook.Save("JapaneseEraDate.pdf", pdfOptions);

            Console.WriteLine("PDF generated with Japanese era date formatting.");
        }
    }
}