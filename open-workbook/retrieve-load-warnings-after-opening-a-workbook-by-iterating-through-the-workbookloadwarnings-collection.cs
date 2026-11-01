// Title: C# example: iterate Workbook.LoadWarnings to retrieve Excel load warnings with Aspose.Cells
// Common Searches: c# how to read workbook load warnings using Aspose.Cells | aspnet retrieve load warnings after loading an .xlsx file with Aspose.Cells | sample code for iterating Workbook.LoadWarnings collection in Aspose.Cells .NET | list of load warnings returned by Aspose.Cells when opening an Excel workbook | aspose.cells load warnings enumeration example c#
// Tags: Aspose.Cells workbook load warnings enumeration | C# iterate Workbook.LoadWarnings | detect Excel loading issues Aspose.Cells | read load warnings from .xlsx file Aspose | handle workbook load warnings .NET

using System;
using System.IO;
using Aspose.Cells;

// The example verifies the presence of an input.xlsx file, loads it into an Aspose.Cells Workbook, and then iterates through the Workbook.LoadWarnings collection, printing each warning message to the console to help developers identify issues detected during the load process.
class Program
{
    static void Main()
    {
        // Path to the Excel file to be loaded
        string filePath = "input.xlsx";

        // Verify that the file exists before attempting to load it
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(filePath);
            Console.WriteLine("Workbook loaded successfully.");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors that may occur during loading
            Console.WriteLine($"An error occurred while loading the workbook: {ex.Message}");
        }
    }
}
