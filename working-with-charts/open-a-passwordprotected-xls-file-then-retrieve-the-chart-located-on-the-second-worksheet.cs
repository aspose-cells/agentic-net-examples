// Title: Open a password‑protected XLS workbook and retrieve the first chart from the second worksheet using Aspose.Cells for .NET
// AI Prompts: Load an encrypted XLS file with Aspose.Cells LoadOptions and obtain the chart object from worksheet index 1 in C#. | Demonstrate how to access the Charts collection of the second worksheet after opening a password‑protected workbook. | Show code to read a protected Excel file, check for charts, and print the chart type using Aspose.Cells.
// Common Searches: aspnet load password protected xls file Aspose.Cells and get chart from second sheet | c# Aspose.Cells retrieve chart from worksheet index 1 in encrypted workbook | how to use LoadOptions.Password to open protected Excel and access chart collection | example code for reading chart from protected XLS using Aspose.Cells for .NET | extract first chart on second worksheet from password secured Excel file C#
// Tags: Aspose.Cells load password protected XLS | retrieve chart from specific worksheet | chart extraction from encrypted workbook | C# access worksheet charts collection | LoadOptions password property usage

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads a password‑protected XLS file via LoadOptions, selects the second worksheet, checks for charts, and retrieves the first chart, outputting its type and parent worksheet name.
class Program
{
    static void Main()
    {
        // Path to the password‑protected XLS file
        string filePath = "protected.xls";
        // Password used to protect the workbook
        string password = "myPassword";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Load the workbook with the password (auto‑detect format)
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = password;
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Access the second worksheet (index 1)
            Worksheet secondWorksheet = workbook.Worksheets[1];

            // Check if the worksheet contains any charts
            if (secondWorksheet.Charts.Count > 0)
            {
                // Retrieve the first chart on the second worksheet
                Chart chart = secondWorksheet.Charts[0];

                // Example usage: display chart type and its worksheet name
                Console.WriteLine("Chart type: " + chart.Type);
                Console.WriteLine("Chart belongs to worksheet: " + chart.Worksheet.Name);
            }
            else
            {
                Console.WriteLine("No charts found on the second worksheet.");
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
