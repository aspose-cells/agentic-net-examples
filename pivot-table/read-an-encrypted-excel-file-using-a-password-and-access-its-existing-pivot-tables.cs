// Title: Read a password‑protected Excel workbook and enumerate its pivot tables with Aspose.Cells for .NET (C#)
// Common Searches: C# load encrypted Excel file using Aspose.Cells LoadOptions password | How to list all pivot tables in a protected workbook with Aspose.Cells | Aspose.Cells read pivot tables from password protected .xlsx in .NET | Iterate worksheets and get pivot table details from a secured Excel file using Aspose.Cells | Example of opening a password‑protected workbook and accessing PivotTable objects in C#
// Tags: LoadOptions password protected workbook Aspose.Cells | Enumerate worksheet pivot tables Aspose.Cells C# | Retrieve pivot table name and source range Aspose.Cells | Open secured .xlsx file Aspose.Cells .NET | Iterate worksheets for pivot tables Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample creates a LoadOptions object with the workbook password, loads the encrypted Excel file into a Workbook, iterates each worksheet, checks for PivotTables, and prints each pivot table’s name, data source range, and destination cell, while handling missing files and exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the encrypted Excel file
            string filePath = "EncryptedWorkbook.xlsx";

            // Password used to protect the workbook
            string password = "YourPassword";

            // Verify that the file exists before attempting to load it
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
                return;
            }

            try
            {
                // Load the workbook with the password using LoadOptions
                LoadOptions loadOptions = new LoadOptions
                {
                    Password = password
                };
                Workbook workbook = new Workbook(filePath, loadOptions);

                // Iterate through all worksheets in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Check if the worksheet contains any pivot tables
                    if (sheet.PivotTables.Count > 0)
                    {
                        Console.WriteLine($"Worksheet: {sheet.Name} contains {sheet.PivotTables.Count} pivot table(s).");

                        // Iterate through each pivot table in the worksheet
                        foreach (var pivotTableObj in sheet.PivotTables)
                        {
                            // Use dynamic to avoid compile‑time dependency on the PivotTable type
                            dynamic pivotTable = pivotTableObj;

                            // Output basic information about the pivot table
                            Console.WriteLine($"  Pivot Table Name: {pivotTable.Name}");
                            Console.WriteLine($"  Data Source Range: {pivotTable.DataSource}");
                            Console.WriteLine($"  Destination Cell: {pivotTable.DestinationCell}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Handle any errors that occur during loading or processing
                Console.WriteLine($"An error occurred: {ex.Message}");
            }

            // Optional: keep console window open
            Console.WriteLine("Processing completed. Press any key to exit.");
            Console.ReadKey();
        }
    }
}
