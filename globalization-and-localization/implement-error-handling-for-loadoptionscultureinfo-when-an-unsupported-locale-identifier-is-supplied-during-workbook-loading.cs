// Title: How to implement fallback to InvariantCulture for unsupported CultureInfo in Aspose.Cells LoadOptions while loading a workbook (C#)
// Common Searches: how to catch CultureNotFoundException when assigning LoadOptions.CultureInfo in Aspose.Cells | fallback to invariant culture for unsupported locale during Excel workbook load with Aspose.Cells | set default culture for Excel import using Aspose.Cells LoadOptions in C# | handle invalid locale identifier with Aspose.Cells LoadOptions CultureInfo | Aspose.Cells workbook loading error handling for custom CultureInfo
// Tags: Aspose.Cells LoadOptions culture fallback | unsupported locale handling workbook load | CultureNotFoundException handling C# | default invariant culture Excel import | globalization error handling Aspose.Cells

using System;
using System.Globalization;
using Aspose.Cells;

// The example demonstrates loading an Excel file with Aspose.Cells while safely assigning a CultureInfo to LoadOptions. It catches CultureNotFoundException for unsupported locale identifiers, falls back to CultureInfo.InvariantCulture, and then proceeds to load and save the workbook.
class WorkbookLoader
{
    static void Main()
    {
        // Path to the Excel file to be loaded
        string filePath = "input.xlsx";

        // Locale identifier supplied by the user or external source
        string localeIdentifier = "invalid-locale"; // Example of an unsupported locale

        // Initialize LoadOptions
        LoadOptions loadOptions = new LoadOptions();

        // Attempt to set CultureInfo based on the supplied locale identifier
        try
        {
            // This will throw CultureNotFoundException if the locale is unsupported
            loadOptions.CultureInfo = new CultureInfo(localeIdentifier);
        }
        catch (CultureNotFoundException ex)
        {
            // Handle unsupported locale identifier
            Console.WriteLine($"Warning: Locale '{localeIdentifier}' is not supported. Using InvariantCulture instead.");
            // Fallback to a safe default culture (InvariantCulture or any other preferred culture)
            loadOptions.CultureInfo = CultureInfo.InvariantCulture;

            // Optional: Log the exception details for diagnostics
            // LogException(ex);
        }

        // Load the workbook with the configured LoadOptions
        Workbook workbook = new Workbook(filePath, loadOptions);

        // (Optional) Perform operations on the workbook here

        // Save the workbook to verify successful loading (optional)
        workbook.Save("output.xlsx");
    }

    // Example placeholder for a logging method
    // static void LogException(Exception ex)
    // {
    //     // Implement logging as needed
    // }
}
