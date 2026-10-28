// Title: Load a password‑protected .xlsx workbook and enumerate its pivot tables with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells open encrypted .xlsx with password and read pivot tables | How to use LoadOptions password to access pivot tables in a protected Excel file | Enumerate pivot tables from a password‑protected workbook using Aspose.Cells .NET | Read pivot table names from a secured Excel file in C# | Aspose.Cells load encrypted workbook and list pivot table details
// Tags: load password protected workbook Aspose.Cells | enumerate pivot tables C# | Aspose.Cells LoadOptions password | pivot table collection iteration Aspose | read encrypted xlsx Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// Demonstrates loading an encrypted .xlsx file by setting the password in LoadOptions, creating a Workbook, iterating through each worksheet to obtain its PivotTableCollection, and printing each pivot table's name, worksheet, and data source while handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        // Path to the encrypted Excel file
        string filePath = "encrypted.xlsx";
        // Password for the encrypted file
        string password = "myPassword";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
            return;
        }

        try
        {
            // Set load options with the password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = password
            };

            // Load the workbook with the specified password
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Iterate through worksheets to find pivot tables
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the collection of pivot tables in the worksheet
                PivotTableCollection pivotTables = sheet.PivotTables;

                // Process each pivot table
                foreach (PivotTable pivot in pivotTables)
                {
                    // Output pivot table details
                    Console.WriteLine($"Worksheet: {sheet.Name}");
                    Console.WriteLine($"Pivot Table Name: {pivot.Name}");
                    Console.WriteLine($"Data Source: {pivot.DataSource}");
                    // Additional operations on the pivot table can be performed here
                }
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
