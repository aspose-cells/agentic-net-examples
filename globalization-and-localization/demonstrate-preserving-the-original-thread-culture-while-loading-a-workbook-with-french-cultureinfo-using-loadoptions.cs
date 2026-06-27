using System;
using System.Globalization;
using System.Threading;
using Aspose.Cells;

class PreserveThreadCultureDemo
{
    static void Main()
    {
        // Store the original thread culture.
        CultureInfo originalCulture = Thread.CurrentThread.CurrentCulture;

        // Change the thread culture to demonstrate that it will be restored later.
        Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

        // Create LoadOptions and set its CultureInfo to French (France).
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.CultureInfo = new CultureInfo("fr-FR");

        // Load the workbook using the specified LoadOptions.
        Workbook workbook = new Workbook("sample.xlsx", loadOptions);

        // Restore the original thread culture after loading.
        Thread.CurrentThread.CurrentCulture = originalCulture;

        // Verify that the workbook uses French culture settings.
        string groupSeparator = workbook.Settings.CultureInfo.NumberFormat.NumberGroupSeparator;
        Console.WriteLine($"French group separator: '{groupSeparator}'");

        // Save the workbook (optional).
        workbook.Save("output.xlsx");
    }
}