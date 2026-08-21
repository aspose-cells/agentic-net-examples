// Title: Load Excel Workbook with French CultureInfo While Preserving Thread Culture – Aspose.Cells C# Example
// Description: C# sample that captures the current thread culture, creates and saves a workbook, then reloads it using LoadOptions.CultureInfo set to "fr-FR". The code demonstrates French number formatting (comma as decimal separator) and verifies that the original thread culture remains unchanged.
// Keywords: Aspose.Cells | LoadOptions | CultureInfo | French culture | fr-FR | thread culture | preserve culture | globalization | localization | C# | Excel loading | regional settings
// Common Searches: Aspose.Cells load workbook with specific CultureInfo | preserve thread culture Aspose.Cells | LoadOptions CultureInfo French example C# | keep original culture when loading Excel with Aspose | Excel number formatting French locale Aspose.Cells
// Developer Intent: Load an Excel file with a designated locale (French) without altering the application's current thread culture.
// Use Cases: Read Excel files created in a French locale while the UI stays in the default language. | Generate reports that require French number formatting without changing the global culture. | Process multi‑regional Excel imports in a single‑threaded service. | Display locale‑specific cell values for auditing while preserving thread settings.
// AI Prompts: Provide a C# code snippet using Aspose.Cells LoadOptions to load an .xlsx with French CultureInfo while keeping Thread.CurrentThread.CurrentCulture unchanged. | Explain the difference between LoadOptions.CultureInfo and the thread's culture in Aspose.Cells. | Show how to verify that the thread culture is unchanged after loading a workbook with a different CultureInfo.

using System;
using System.Globalization;
using System.Threading;
using Aspose.Cells;

// C# sample that captures the current thread culture, creates and saves a workbook, then reloads it using LoadOptions.CultureInfo set to "fr-FR". The code demonstrates French number formatting (comma as decimal separator) and verifies that the original thread culture remains unchanged.
class Program
{
    static void Main()
    {
        // Preserve the original thread culture
        CultureInfo originalCulture = Thread.CurrentThread.CurrentCulture;
        Console.WriteLine("Original thread culture: " + originalCulture.Name);

        // -------------------------------------------------
        // Create a sample workbook (using the standard constructor)
        // -------------------------------------------------
        Workbook sampleWorkbook = new Workbook();
        // Put a numeric value to see culture‑specific formatting later
        sampleWorkbook.Worksheets[0].Cells["A1"].PutValue(1234.56);
        // Save the workbook (using the standard Save method)
        sampleWorkbook.Save("sample.xlsx");

        // -------------------------------------------------
        // Load the workbook with French regional settings
        // -------------------------------------------------
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        // Set the CultureInfo to French (France)
        loadOptions.CultureInfo = new CultureInfo("fr-FR");

        // Load the workbook with the specified LoadOptions
        Workbook loadedWorkbook = new Workbook("sample.xlsx", loadOptions);

        // Display the cell value as a string; French culture uses comma as decimal separator
        Console.WriteLine("Loaded cell string value (French culture): " + loadedWorkbook.Worksheets[0].Cells["A1"].StringValue);

        // -------------------------------------------------
        // Verify that the thread culture has not changed
        // -------------------------------------------------
        Console.WriteLine("Thread culture after load: " + Thread.CurrentThread.CurrentCulture.Name);
        Console.WriteLine("Thread culture unchanged: " + originalCulture.Equals(Thread.CurrentThread.CurrentCulture));
    }
}
