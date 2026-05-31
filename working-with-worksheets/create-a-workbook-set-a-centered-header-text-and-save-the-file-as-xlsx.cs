using System;
using Aspose.Cells;

namespace AsposeCellsHeaderExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses Workbook() constructor rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the PageSetup object for the worksheet
            PageSetup pageSetup = worksheet.PageSetup;

            // Set the centered header (section 1 = center) with desired text
            // Uses PageSetup.SetHeader(int section, string headerScript) rule
            pageSetup.SetHeader(1, "My Centered Header");

            // Save the workbook as XLSX (uses Workbook.Save(string) rule)
            workbook.Save("CenteredHeaderWorkbook.xlsx");

            // Optional: inform the user
            Console.WriteLine("Workbook created with centered header and saved as 'CenteredHeaderWorkbook.xlsx'.");
        }
    }
}