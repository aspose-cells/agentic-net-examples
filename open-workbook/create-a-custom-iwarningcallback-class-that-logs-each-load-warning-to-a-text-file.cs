// Title: Log workbook load warnings to a text file with a custom IWarningCallback in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# load options warningcallback write warnings to log file | how to capture Excel load warnings using IWarningCallback in .NET | save Aspose.Cells workbook load warnings to a custom log file | implement IWarningCallback to log warningInfo details during workbook load | C# example of logging Aspose.Cells load warnings with timestamp
// Tags: custom IWarningCallback for Aspose.Cells | logging workbook load warnings to file | Aspose.Cells LoadOptions warning callback | C# Aspose.Cells warning logging example | Excel import warning handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// This example demonstrates creating a FileWarningCallback class that implements Aspose.Cells.IWarningCallback, ensures the log directory exists, and appends each warning (UTC timestamp, type, description) to a specified text file. The callback is assigned to LoadOptions.WarningCallback, enabling automatic logging of any load warnings when a Workbook is opened, after which the workbook can be processed or saved.
public class FileWarningCallback : IWarningCallback
{
    private readonly string _logFilePath;

    public FileWarningCallback(string logFilePath)
    {
        _logFilePath = logFilePath;

        // Ensure the directory for the log file exists
        string directory = Path.GetDirectoryName(_logFilePath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }
    }

    // Implementation required by IWarningCallback
    public void Warning(WarningInfo warningInfo)
    {
        try
        {
            // Build a log entry with timestamp, warning type and description
            string logEntry = $"{DateTime.UtcNow:O} | Type: {warningInfo.WarningType} | Description: {warningInfo.Description}";

            // Append the entry to the log file
            File.AppendAllText(_logFilePath, logEntry + Environment.NewLine);
        }
        catch (Exception ex)
        {
            // If logging fails, write to console (avoid throwing from callback)
            Console.WriteLine($"Failed to log warning: {ex.Message}");
        }
    }
}

class Program
{
    static void Main()
    {
        try
        {
            // Path to the Excel file to be loaded
            string excelFilePath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(excelFilePath))
            {
                Console.WriteLine($"Error: The file '{excelFilePath}' does not exist.");
                return;
            }

            // Path where warnings will be logged
            string warningLogPath = "warnings.log";

            // Set up load options and assign the custom warning callback
            LoadOptions loadOptions = new LoadOptions
            {
                WarningCallback = new FileWarningCallback(warningLogPath)
            };

            // Load the workbook; any warnings during loading will be logged
            Workbook workbook = new Workbook(excelFilePath, loadOptions);

            // Example: save the workbook to a new file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log unexpected exceptions
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
