// Title: Prompt user to select a new external data source file and update an Excel workbook’s DataConnection using Aspose.Cells for .NET
// AI Prompts: Generate a C# console application that opens a given workbook, shows a file‑picker prompt for the user to choose a replacement external data source, assigns the chosen path to the first DataConnection, and saves the workbook under a new name. | Write C# code with Aspose.Cells that verifies the existence of both the input workbook and the selected source file, updates the workbook’s external connection to point to the new file, and includes robust error handling.
// Common Searches: how to use Aspose.Cells to change the source file of an external data connection in an Excel workbook | C# console program to ask user for a file path and update Excel external connection | Aspose.Cells .NET update DataConnection source after user selects file | validate workbook and external source file existence before modifying connection Aspose.Cells | save workbook with updated external connection using Aspose.Cells C#
// Tags: modify external connection source Aspose.Cells | C# console file selection for Excel data connection | check workbook and source file existence .NET | save workbook after external connection update Aspose.Cells | replace external data source path in Excel using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

// // Loads an existing workbook, prompts the user via a console file picker to choose a new external data source, updates the first DataConnection's SourceFile property, and saves the modified workbook as output.xlsx.
class Program
{
    // Entry point of the console application.
    static void Main()
    {
        try
        {
            // Path to the workbook that contains an external connection.
            const string inputWorkbookPath = "input.xlsx";

            // Verify that the input workbook exists.
            if (!File.Exists(inputWorkbookPath))
            {
                Console.WriteLine($"Error: The workbook \"{inputWorkbookPath}\" was not found.");
                return;
            }

            // Prompt the user to provide the external data source file path.
            Console.Write("Enter the full path of the external data source file: ");
            string sourceFilePath = Console.ReadLine()?.Trim();

            // Validate the provided source file path.
            if (string.IsNullOrEmpty(sourceFilePath) || !File.Exists(sourceFilePath))
            {
                Console.WriteLine("Error: The specified external data source file does not exist.");
                return;
            }

            // Load the workbook.
            Workbook workbook = new Workbook(inputWorkbookPath);

            // Ensure the workbook contains at least one external connection.
            if (workbook.DataConnections.Count == 0)
            {
                Console.WriteLine("Error: No external connections found in the workbook.");
                return;
            }

            // Update the first external connection's source file.
            ExternalConnection connection = workbook.DataConnections[0];
            connection.SourceFile = sourceFilePath;

            // Save the workbook with the updated connection.
            const string outputWorkbookPath = "output.xlsx";
            workbook.Save(outputWorkbookPath);

            Console.WriteLine($"Success: External data source updated and workbook saved as \"{outputWorkbookPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message.
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
