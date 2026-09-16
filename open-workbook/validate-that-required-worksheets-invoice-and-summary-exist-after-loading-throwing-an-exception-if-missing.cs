// Title: C# – Validate that the 'Invoice' and 'Summary' worksheets exist in an Aspose.Cells workbook after loading
// AI Prompts: Generate C# code that loads an Excel file with Aspose.Cells and throws an InvalidOperationException when either the 'Invoice' or 'Summary' worksheet is missing. | Create a reusable C# helper method using Aspose.Cells that receives a Workbook and a list of sheet names, checks each for presence, and raises an exception for any absent sheet.
// Common Searches: asp.net load excel with Aspose.Cells and ensure Invoice sheet exists | c# Aspose.Cells validate presence of Summary worksheet after opening workbook | how to throw error if required worksheet not found in Aspose.Cells workbook | check multiple required worksheets in an Aspose.Cells workbook using C#
// Tags: Aspose.Cells worksheet existence check | C# validate required Excel sheets | Aspose.Cells InvalidOperationException missing sheet | Excel workbook required sheet verification C# | Aspose.Cells load workbook with default sheets

using Aspose.Cells;
using System;
using System.IO;

// The program loads an Excel file using Aspose.Cells, creates it with 'Invoice' and 'Summary' sheets if it does not exist, and then validates that both worksheets are present, throwing an InvalidOperationException if either is missing.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the Excel file to be loaded
            string filePath = "input.xlsx";

            Workbook workbook;

            // Ensure the input file exists; if not, create a new workbook with required sheets
            if (File.Exists(filePath))
            {
                // Load the workbook from the specified file
                workbook = new Workbook(filePath);
            }
            else
            {
                // Create a new workbook and add the required worksheets
                workbook = new Workbook();
                workbook.Worksheets.Clear(); // Remove the default sheet
                workbook.Worksheets.Add("Invoice");
                workbook.Worksheets.Add("Summary");
                // Optionally save the newly created workbook for future runs
                workbook.Save(filePath);
                Console.WriteLine($"Input file not found. A new workbook with required sheets has been created at \"{filePath}\".");
            }

            // Validate that the required worksheets exist
            ValidateWorksheets(workbook, new[] { "Invoice", "Summary" });

            // Further processing can continue here
            Console.WriteLine("All required worksheets are present.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Checks for the presence of each required worksheet name
    static void ValidateWorksheets(Workbook workbook, string[] requiredSheets)
    {
        foreach (string sheetName in requiredSheets)
        {
            // Attempt to retrieve the worksheet by name
            Worksheet sheet = workbook.Worksheets[sheetName];

            // If the worksheet is not found, throw an exception
            if (sheet == null)
            {
                throw new InvalidOperationException($"Required worksheet \"{sheetName}\" is missing.");
            }
        }
    }
}
