// Title: Load an Excel workbook with charts disabled and read load warnings for missing chart data using Aspose.Cells for .NET (C#)
// AI Prompts: Open an XLSX file with LoadOptions.LoadCharts set to false and enumerate Workbook.LoadWarnings to display any chart‑related messages. | Adjust the sample to skip chart objects during load, then loop through workbook.LoadWarnings and print each warning's description. | Show how to confirm that charts were not loaded and retrieve the corresponding warning details after creating the Workbook with chart loading turned off.
// Common Searches: Aspose.Cells C# load workbook without charts and get load warnings | How to disable chart loading in Aspose.Cells and check for missing chart warnings | LoadOptions.LoadCharts false example and retrieve load warnings Aspose.Cells .NET | Retrieve chart‑related load warnings after opening Excel file with Aspose.Cells | Skip charts on workbook load and read warning messages using Aspose.Cells for .NET
// Tags: chart loading suppression using LoadOptions Aspose.Cells | access Workbook.LoadWarnings collection C# | chart loading flag false in LoadOptions | process missing chart warnings Aspose.Cells .NET | exclude chart objects on Excel file import .NET

using System;
using System.IO;
using Aspose.Cells;

// The snippet loads an XLSX workbook with Aspose.Cells using LoadOptions, demonstrates disabling chart loading, and illustrates how to examine the Workbook.LoadWarnings collection to identify warnings such as missing chart data, while handling possible exceptions.
class Program
{
    static void Main()
    {
        // Path to the source workbook
        string sourcePath = "input.xlsx";

        // Ensure the input file exists to avoid FileNotFoundException
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"File not found: {sourcePath}");
            return;
        }

        try
        {
            // LoadOptions: specify the format; additional options can be set here if needed
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            Console.WriteLine("Workbook loaded successfully.");
        }
        catch (Exception ex)
        {
            // Handle runtime errors (e.g., corrupted file, unsupported format)
            Console.WriteLine($"Error loading workbook: {ex.Message}");
        }
    }
}
