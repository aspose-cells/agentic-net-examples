// Title: Load an XLSX workbook with French (fr-FR) CultureInfo using Aspose.Cells LoadOptions in C#
// Description: Shows how to apply a French CultureInfo to a LoadOptions object, load an XLSX file with Aspose.Cells, read a cell value, and then restore the original thread culture so the rest of the application remains unaffected.
// Keywords: Aspose.Cells LoadOptions CultureInfo | C# load Excel French locale | preserve thread culture | globalization Aspose.Cells | fr-FR Excel parsing | Aspose.Cells localization example | Excel date format France | LoadOptions example C#
// Common Searches: Aspose.Cells load workbook with French culture | Set CultureInfo for Excel load in .NET | Keep original thread culture when using LoadOptions | C# Aspose.Cells LoadOptions CultureInfo sample | Parse French formatted dates in Excel with Aspose.Cells
// Developer Intent: Load an XLSX file using a French locale while leaving the application's thread culture unchanged.
// Use Cases: Read French‑formatted dates and numbers from an Excel sheet without altering UI culture. | Generate localized reports from French workbooks on a server that must retain its default culture for other tasks. | Process multiple workbooks in parallel, each with its own CultureInfo, by resetting the thread culture after each load.
// AI Prompts: Provide C# code to load an Excel workbook with German (de-DE) CultureInfo using Aspose.Cells LoadOptions and restore the original thread culture. | Explain how LoadOptions.CultureInfo affects number, date, and currency parsing in Aspose.Cells and the best practice for safely reverting thread culture. | Show a pattern for loading several workbooks concurrently, assigning a distinct CultureInfo to each LoadOptions instance.

using System;
using System.Globalization;
using System.Threading;
using Aspose.Cells;

// Shows how to apply a French CultureInfo to a LoadOptions object, load an XLSX file with Aspose.Cells, read a cell value, and then restore the original thread culture so the rest of the application remains unaffected.
class Program
{
    static void Main()
    {
        // Preserve the original thread culture
        CultureInfo originalCulture = Thread.CurrentThread.CurrentCulture;

        // Create LoadOptions and set the culture to French (France)
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.CultureInfo = new CultureInfo("fr-FR");

        // Load the workbook using the constructor that accepts LoadOptions
        Workbook workbook = new Workbook("sample.xlsx", loadOptions);

        // Example operation: output the value of cell A1
        Console.WriteLine("A1 value: " + workbook.Worksheets[0].Cells["A1"].StringValue);

        // Restore the original thread culture
        Thread.CurrentThread.CurrentCulture = originalCulture;
    }
}
