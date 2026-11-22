// Title: Load an XLSX workbook in C# with French (fr-FR) locale using Aspose.Cells LoadOptions while preserving the original thread culture
// Common Searches: how to load an XLSX file with French (fr-FR) culture using Aspose.Cells in C# | preserve original thread culture when changing CultureInfo for Aspose.Cells workbook loading | Aspose.Cells LoadOptions respecting thread CultureInfo for number and date parsing | set thread CultureInfo to French before loading Excel workbook with Aspose.Cells | restore previous CultureInfo after loading workbook with Aspose.Cells LoadOptions
// Tags: Aspose.Cells LoadOptions with thread CultureInfo | load XLSX workbook using French fr-FR locale | temporary CultureInfo change for Excel parsing in C# | restore original CultureInfo after Aspose.Cells load | culture-specific number and date parsing Aspose.Cells

using System;
using System.Globalization;
using Aspose.Cells;

// The example temporarily switches the thread's CultureInfo to French (fr-FR), creates LoadOptions for XLSX (which inherits the thread culture), loads the workbook with those options, and finally restores the original thread culture.
class Program
{
    static void Main()
    {
        // Preserve the original thread culture
        CultureInfo originalCulture = CultureInfo.CurrentCulture;
        CultureInfo originalUICulture = CultureInfo.CurrentUICulture;

        try
        {
            // Set the thread culture to French (France)
            CultureInfo frenchCulture = new CultureInfo("fr-FR");
            CultureInfo.CurrentCulture = frenchCulture;
            CultureInfo.CurrentUICulture = frenchCulture;

            // Create LoadOptions for XLSX format (culture is taken from the thread)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // Load the workbook using the specified LoadOptions
            Workbook workbook = new Workbook("input.xlsx", loadOptions);

            // Workbook is now loaded with French culture settings (e.g., number/date parsing)
            // Perform any required operations on the workbook here
        }
        finally
        {
            // Restore the original thread culture
            CultureInfo.CurrentCulture = originalCulture;
            CultureInfo.CurrentUICulture = originalUICulture;
        }
    }
}
