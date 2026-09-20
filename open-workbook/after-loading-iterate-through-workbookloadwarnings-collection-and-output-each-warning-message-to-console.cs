// Title: Read and display Aspose.Cells Workbook LoadWarnings messages in a C# console application
// AI Prompts: Generate C# code that opens an Excel file with Aspose.Cells, checks the Workbook.LoadWarnings collection, and writes each warning's Message to the console. | Create a C# example that uses reflection to safely access the LoadWarnings property of a Workbook and prints all warning texts. | Write a C# console program that loads a spreadsheet, handles missing‑file errors, and enumerates any load warnings returned by Aspose.Cells.
// Common Searches: how to enumerate load warnings after loading an Excel workbook with Aspose.Cells in C# | C# Aspose.Cells get warning messages when opening a corrupted xlsx file | sample code for reading Workbook.LoadWarnings property using reflection in .NET | display Aspose.Cells load warnings in a console application
// Tags: Aspose.Cells Workbook.LoadWarnings enumeration | C# console output of Excel load warnings | reflection access to Aspose.Cells LoadWarnings property | handling missing Excel file with Aspose.Cells | exception handling for workbook loading in .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook using Aspose.Cells, verifies the file exists, optionally accesses the LoadWarnings property via reflection, iterates through any warnings, and prints each warning's Message to the console while handling errors gracefully.
class Program
{
    static void Main()
    {
        const string filePath = "input.xlsx";

        // Ensure the input file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(filePath);

            // Attempt to retrieve load warnings if the API version supports it
            var loadWarningsProp = typeof(Workbook).GetProperty("LoadWarnings");
            if (loadWarningsProp != null)
            {
                var warnings = loadWarningsProp.GetValue(workbook) as System.Collections.IEnumerable;
                if (warnings != null)
                {
                    foreach (var warning in warnings)
                    {
                        // Each warning is expected to have a Message property
                        var messageProp = warning.GetType().GetProperty("Message");
                        if (messageProp != null)
                        {
                            string message = messageProp.GetValue(warning) as string;
                            Console.WriteLine(message);
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // Catch any runtime errors (e.g., loading issues) and display a friendly message
            Console.WriteLine($"Error loading workbook: {ex.Message}");
        }
    }
}
