using System;
using Aspose.Cells;

namespace AsposeCellsPaperSizeDemo
{
    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set a predefined paper size (A4)
            sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;

            // Set a custom paper size (width: 5 inches, height: 7 inches)
            sheet.PageSetup.CustomPaperSize(5.0, 7.0);

            // Save the workbook to a file
            workbook.Save("PaperSizeDemo.xlsx");

            // ---------- Load the workbook with a specific default paper size ----------
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.SetPaperSize(PaperSizeType.PaperLetter); // default to Letter size when loading

            Workbook loadedWorkbook = new Workbook("PaperSizeDemo.xlsx", loadOptions);

            // Output the paper size of the loaded workbook's first worksheet
            Console.WriteLine("Loaded workbook paper size: " + loadedWorkbook.Worksheets[0].PageSetup.PaperSize);
        }
    }
}