using System;
using System.Globalization;
using System.Threading;
using Aspose.Cells;

class PreserveThreadCultureDemo
{
    static void Main()
    {
        // Store the original thread culture (e.g., en-US)
        CultureInfo originalCulture = Thread.CurrentThread.CurrentCulture;

        // Create a sample workbook with a numeric value
        Workbook sample = new Workbook();
        sample.Worksheets[0].Cells["A1"].PutValue(1234.56);
        // Save the sample workbook to a file
        sample.Save("sample.xlsx");

        // Ensure the current thread culture is still the original one
        Console.WriteLine("Thread culture before loading: " + Thread.CurrentThread.CurrentCulture.Name);

        // Create LoadOptions and set its CultureInfo to French (fr-FR)
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.CultureInfo = new CultureInfo("fr-FR"); // French uses comma as decimal separator

        // Load the workbook using the specified LoadOptions
        Workbook loaded = new Workbook("sample.xlsx", loadOptions);

        // Retrieve the value of cell A1 as a string; it will be formatted according to French culture
        string frenchFormattedValue = loaded.Worksheets[0].Cells["A1"].StringValue;
        Console.WriteLine("Value formatted with French culture: " + frenchFormattedValue);

        // Verify that the thread culture has not changed after loading
        Console.WriteLine("Thread culture after loading: " + Thread.CurrentThread.CurrentCulture.Name);

        // Restore the original culture (good practice if it was changed elsewhere)
        Thread.CurrentThread.CurrentCulture = originalCulture;
    }
}