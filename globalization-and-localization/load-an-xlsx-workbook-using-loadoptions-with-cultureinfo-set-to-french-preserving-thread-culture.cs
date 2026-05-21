using System;
using System.Globalization;
using System.Threading;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Preserve the original thread culture
        CultureInfo originalCulture = Thread.CurrentThread.CurrentCulture;

        try
        {
            // Set thread culture to French (France) to match the load options
            Thread.CurrentThread.CurrentCulture = new CultureInfo("fr-FR");

            // Create LoadOptions and assign French culture
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.CultureInfo = new CultureInfo("fr-FR");

            // Load the XLSX workbook using the specified LoadOptions
            Workbook workbook = new Workbook("input.xlsx", loadOptions);

            // Example usage: read the string value of cell A1
            string cellValue = workbook.Worksheets[0].Cells["A1"].StringValue;
            Console.WriteLine($"A1 value: {cellValue}");
        }
        finally
        {
            // Restore the original thread culture
            Thread.CurrentThread.CurrentCulture = originalCulture;
        }
    }
}