// Title: Load an Excel workbook with Aspose.Cells in C# and list its ContentTypeProperties
// Common Searches: how to load an existing .xlsx file into Aspose.Cells Workbook using C# | C# code to enumerate ContentTypeProperties of an Excel workbook with Aspose.Cells | retrieve custom content type metadata from a workbook using Aspose.Cells .NET | example of handling FileNotFoundException when opening Excel file with Aspose.Cells | display workbook content type properties in console using Aspose.Cells C#
// Tags: load workbook from file Aspose.Cells C# | list workbook content type properties Aspose.Cells | read custom content type metadata Excel | handle missing workbook file C# | display workbook properties in console Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example checks for the presence of 'input.xlsx', loads it into an Aspose.Cells Workbook, accesses the ContentTypeProperties collection, and prints each property's name and value to the console while gracefully handling missing files and other exceptions.
class Program
{
    static void Main()
    {
        const string filePath = "input.xlsx";

        try
        {
            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(filePath);

            // Retrieve the collection of content type properties
            var properties = workbook.ContentTypeProperties;

            // Iterate and display each property's name and value
            foreach (var prop in properties)
            {
                Console.WriteLine($"{prop.Name}: {prop.Value}");
            }
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
