// Title: Load an XLSX workbook with French (fr-FR) CultureInfo using Aspose.Cells LoadOptions while preserving original thread cultures in C#
// AI Prompts: Use Aspose.Cells LoadOptions to load a .xlsx file with the French locale (fr-FR) in C#, then restore the thread's original CultureInfo settings. | Show how to temporarily set Thread.CurrentThread culture to French for workbook loading with Aspose.Cells and revert to the saved cultures after the load.
// Common Searches: C# Aspose.Cells load xlsx with French CultureInfo without affecting global thread culture | How to set LoadOptions.CultureInfo to fr-FR when opening an Excel file in Aspose.Cells | Preserve original thread culture while changing CultureInfo for Aspose.Cells workbook loading | Aspose.Cells LoadOptions French locale example for .xlsx files in .NET
// Tags: Aspose.Cells LoadOptions CultureInfo French | C# load XLSX with specific locale | preserve thread culture Aspose.Cells | globalization Excel parsing Aspose.Cells | temporary thread culture change .NET

using System;
using System.Globalization;
using System.Threading;
using Aspose.Cells;

// The program saves the current thread's CultureInfo, switches the thread to French (fr-FR), creates LoadOptions with that CultureInfo, loads an XLSX workbook using those options, and finally restores the original thread cultures.
class Program
{
    static void Main()
    {
        // Preserve the original thread cultures
        CultureInfo originalCulture = Thread.CurrentThread.CurrentCulture;
        CultureInfo originalUICulture = Thread.CurrentThread.CurrentUICulture;

        // Set thread culture to French (France) for loading
        CultureInfo frenchCulture = new CultureInfo("fr-FR");
        Thread.CurrentThread.CurrentCulture = frenchCulture;
        Thread.CurrentThread.CurrentUICulture = frenchCulture;

        // Create LoadOptions and assign the French culture
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        loadOptions.CultureInfo = frenchCulture; // ensures parsing respects French locale

        // Load the workbook using the specified LoadOptions
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // TODO: work with the workbook as needed

        // Restore the original thread cultures
        Thread.CurrentThread.CurrentCulture = originalCulture;
        Thread.CurrentThread.CurrentUICulture = originalUICulture;
    }
}
