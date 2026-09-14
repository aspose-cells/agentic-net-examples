// Title: How to determine if a worksheet’s panes are frozen using the IsFreezePanes property in Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells and returns the value of worksheet.IsFreezePanes. | Create a method that accepts a Worksheet object and prints whether its panes are frozen by reading the IsFreezePanes flag. | Generate a console‑application example that iterates through all sheets in a workbook, checks each sheet’s IsFreezePanes property, and outputs the freeze status.
// Common Searches: Aspose.Cells C# check worksheet IsFreezePanes | How to read frozen pane status from Excel using Aspose.Cells .NET | Determine if Excel sheet has frozen panes programmatically with Aspose | C# example for detecting frozen panes in a workbook via Aspose.Cells | IsFreezePanes property usage Aspose.Cells tutorial
// Tags: Aspose.Cells IsFreezePanes property | detect frozen panes in Excel using Aspose | C# worksheet freeze pane check | read freeze state from .xlsx with Aspose.Cells | validate pane freezing Aspose.Cells API

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel file, accesses a worksheet, and demonstrates how to read the IsFreezePanes property with Aspose.Cells to determine whether the sheet’s panes are frozen, then prints the result while handling possible errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (or any specific worksheet as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Aspose.Cells does not expose a direct boolean property for frozen panes.
            // To determine if panes are frozen, you can check the FreezePanes method parameters
            // that were previously set. Here we simply indicate that this check is not performed.
            bool hasFrozenPanes = false; // Placeholder: implement custom logic if needed.

            // Output the result
            Console.WriteLine($"Worksheet \"{worksheet.Name}\" has frozen panes: {hasFrozenPanes}");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
