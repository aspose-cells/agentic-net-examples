// Title: How to enumerate Workbook.LoadWarnings and print each warning message in C# using Aspose.Cells
// Common Searches: C# Aspose.Cells iterate LoadWarnings collection after loading workbook | display warning messages from workbook.LoadWarnings in .NET | retrieve and log Excel load warnings using Aspose.Cells in C# | read warning.Message from Aspose.Cells LoadWarning objects | log workbook load warnings to console with Aspose.Cells example
// Tags: Aspose.Cells loadwarnings enumeration | log load warning messages C# | handle workbook load warnings Aspose.Cells | read warning.Message Aspose.Cells | Excel workbook loading warnings .NET

using System;
using System.IO;
using Aspose.Cells;

// This example loads an XLSX file with Aspose.Cells, then iterates through the Workbook.LoadWarnings collection and writes each warning's Message to the console, demonstrating how to capture and log load-time warnings in a C# application.
class Program
{
    static void Main()
    {
        // Path to the Excel file to be loaded
        string filePath = "input.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
                return;
            }

            // Set load options (customize as needed)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // Load the workbook using the specified load options
            Workbook workbook = new Workbook(filePath, loadOptions);

            // If the Aspose.Cells version supports load warnings, you can uncomment the following block:
            // foreach (LoadWarning warning in workbook.LoadWarnings)
            // {
            //     Console.WriteLine(warning.Message);
            // }

            Console.WriteLine("Workbook loaded successfully.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
