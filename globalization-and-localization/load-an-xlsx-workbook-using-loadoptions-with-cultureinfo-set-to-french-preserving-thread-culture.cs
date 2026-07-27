using System;
using System.Globalization;
using System.Threading;
using Aspose.Cells;

class LoadWorkbookWithFrenchCulture
{
    static void Main()
    {
        // Preserve the original thread culture
        CultureInfo originalCulture = Thread.CurrentThread.CurrentCulture;

        try
        {
            // Create LoadOptions and set French culture (fr-FR)
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.CultureInfo = new CultureInfo("fr-FR");

            // Load the workbook using the constructor that accepts file path and LoadOptions
            Workbook workbook = new Workbook("input.xlsx", loadOptions);

            // Example: read a cell value formatted according to French culture
            Worksheet sheet = workbook.Worksheets[0];
            Console.WriteLine("A1 value (French culture): " + sheet.Cells["A1"].StringValue);
        }
        finally
        {
            // Restore the original thread culture
            Thread.CurrentThread.CurrentCulture = originalCulture;
        }
    }
}