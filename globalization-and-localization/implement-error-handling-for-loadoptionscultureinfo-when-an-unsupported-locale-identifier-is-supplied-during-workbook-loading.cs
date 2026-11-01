// Title: Handle unsupported locale identifiers with LoadOptions.CultureInfo and fallback to InvariantCulture in Aspose.Cells for .NET
// Common Searches: Aspose.Cells LoadOptions CultureInfo invalid locale handling in C# | fallback to InvariantCulture when loading Excel workbook with unsupported culture | catch CultureNotFoundException while setting LoadOptions.CultureInfo Aspose.Cells | set default culture for workbook loading in Aspose.Cells .NET | how to provide safe locale for Aspose.Cells workbook load options
// Tags: LoadOptions CultureInfo fallback | unsupported locale handling Aspose.Cells | CultureNotFoundException Aspose.Cells .NET | InvariantCulture workbook loading | globalization settings Aspose.Cells

using System;
using System.Globalization;
using Aspose.Cells;

// Demonstrates assigning a user‑supplied locale to LoadOptions.CultureInfo, catching CultureNotFoundException for unsupported identifiers, switching to CultureInfo.InvariantCulture, and then loading and saving the workbook with the safe culture settings.
class WorkbookLoader
{
    static void Main()
    {
        // Path to the Excel file to be loaded
        string inputPath = "Sample.xlsx";

        // Locale identifier supplied by the user or external source
        string localeIdentifier = "invalid-locale"; // Example of an unsupported locale

        // Prepare LoadOptions with culture handling
        LoadOptions loadOptions = new LoadOptions();

        try
        {
            // Attempt to assign the requested culture
            loadOptions.CultureInfo = new CultureInfo(localeIdentifier);
        }
        catch (CultureNotFoundException)
        {
            // The supplied locale identifier is not supported.
            // Fallback to InvariantCulture (or any default you prefer)
            Console.WriteLine($"Locale '{localeIdentifier}' is not supported. Using InvariantCulture instead.");
            loadOptions.CultureInfo = CultureInfo.InvariantCulture;
        }

        // Load the workbook using the configured LoadOptions
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // (Optional) Perform any workbook operations here

        // Save the workbook to verify successful loading
        string outputPath = "LoadedOutput.xlsx";
        workbook.Save(outputPath);
    }
}
