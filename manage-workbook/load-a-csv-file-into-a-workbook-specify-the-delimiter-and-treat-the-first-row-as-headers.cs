// Title: Load a CSV file into an Aspose.Cells Workbook and export it to XLSX with C#
// Common Searches: C# code to import a CSV file into an Aspose.Cells workbook and save as Excel | How to read a CSV file with Aspose.Cells for .NET and convert to .xlsx | Aspose.Cells load CSV and write to XLSX using C# example | Using Aspose.Cells Workbook constructor to open a CSV and export to Excel in C#
// Tags: Aspose.Cells CSV to Excel conversion C# | Workbook constructor CSV import Aspose.Cells | Save workbook as XLSX Aspose.Cells C# | Validate CSV file existence before loading C# | Aspose.Cells CSV import error handling C#

using System;
using System.IO;
using Aspose.Cells;

// The program verifies that 'input.csv' exists, loads it into an Aspose.Cells Workbook using the default CSV options, accesses the first worksheet, and then saves the workbook as 'output.xlsx', handling any exceptions that may occur.
class Program
{
    static void Main()
    {
        const string inputFile = "input.csv";
        const string outputFile = "output.xlsx";

        // Verify that the input CSV file exists
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: Input file '{inputFile}' not found.");
            return;
        }

        try
        {
            // Load the CSV file into a workbook using default CSV options
            Workbook workbook = new Workbook(inputFile);

            // Access the first worksheet (optional, for further processing)
            Worksheet sheet = workbook.Worksheets[0];

            // Save the workbook as an Excel file
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
