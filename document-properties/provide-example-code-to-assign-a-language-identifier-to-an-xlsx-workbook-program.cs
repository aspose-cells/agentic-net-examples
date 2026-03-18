using System;
using Aspose.Cells;

namespace AsposeCellsLanguageDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses the provided creation rule)
            Workbook workbook = new Workbook();

            // Access the workbook settings
            WorkbookSettings settings = workbook.Settings;

            // Assign a language identifier (CountryCode) to the workbook
            // This uses the WorkbookSettings.LanguageCode property rule
            settings.LanguageCode = CountryCode.Germany;

            // Optional: display the assigned language code
            Console.WriteLine("Workbook language code set to: " + settings.LanguageCode);

            // Save the workbook to an XLSX file (uses the provided save rule)
            workbook.Save("Workbook_With_LanguageCode.xlsx");
        }
    }
}