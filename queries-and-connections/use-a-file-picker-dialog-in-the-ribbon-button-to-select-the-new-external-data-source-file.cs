using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class Program
{
    static void Main()
    {
        const string inputPath = "InputWithConnection.xlsx";
        const string outputPath = "OutputWithUpdatedConnection.xlsx";

        // Verify the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Ensure there is at least one external connection
            if (workbook.DataConnections.Count == 0)
            {
                Console.WriteLine("Error: No external connections found in the workbook.");
                return;
            }

            ExternalConnection connection = workbook.DataConnections[0];

            // Prompt user for new source file path
            Console.WriteLine("Enter the full path of the new external data source file:");
            string newSourcePath = Console.ReadLine()?.Trim();

            if (string.IsNullOrEmpty(newSourcePath) || !File.Exists(newSourcePath))
            {
                Console.WriteLine("Error: Specified source file does not exist. The external connection remains unchanged.");
                return;
            }

            // Update the connection's source file
            connection.SourceFile = newSourcePath;
            Console.WriteLine($"External connection source file set to: {connection.SourceFile}");

            // Save the workbook with updated connection information
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}