// Title: C# – Open a password‑protected Excel workbook and enumerate its pivot tables using Aspose.Cells
// Description: Demonstrates how to load an encrypted .xlsx file with a password via LoadOptions, iterate through each worksheet, detect existing pivot tables, output their names and row‑field counts, refresh and recalculate the pivot data, and handle missing‑file, invalid‑password and runtime exceptions.
// Keywords: Aspose.Cells encrypted workbook | C# open password protected Excel | enumerate pivot tables Aspose.Cells | refresh pivot cache Aspose.Cells | LoadOptions password Excel .NET | handle CellsException | pivot table calculation C#
// Common Searches: open password protected Excel file Aspose.Cells C# | list pivot tables in encrypted workbook using Aspose.Cells | refresh pivot tables after loading protected Excel file | Aspose.Cells invalid password exception handling | C# code to read encrypted .xlsx and access pivot tables
// Developer Intent: Load a password‑protected Excel file and work with its existing pivot tables programmatically.
// Use Cases: Read a secured workbook and display the count of pivot tables per sheet. | Retrieve each pivot table’s name and row‑field count for reporting. | Refresh the pivot cache and recalculate data after opening the protected file. | Gracefully capture errors such as missing file, wrong password, or pivot refresh failures.
// AI Prompts: Write C# code that opens a password‑protected Excel workbook with Aspose.Cells and lists all pivot table names. | Show how to refresh and calculate pivot tables in an encrypted workbook, including exception handling for invalid passwords. | Explain the steps to catch CellsException when loading a protected workbook and to log detailed error information.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsEncryptedPivotDemo
{
    // Demonstrates how to load an encrypted .xlsx file with a password via LoadOptions, iterate through each worksheet, detect existing pivot tables, output their names and row‑field counts, refresh and recalculate the pivot data, and handle missing‑file, invalid‑password and runtime exceptions.
    class Program
    {
        static void Main()
        {
            // Path to the encrypted Excel file
            string filePath = "encrypted.xlsx";

            // Password used to protect the workbook
            string password = "xixi";

            // Verify that the file exists before attempting to load it
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
                return;
            }

            try
            {
                // Set load options with the password for the encrypted workbook
                LoadOptions loadOptions = new LoadOptions
                {
                    Password = password
                };

                // Load the password‑protected workbook
                Workbook workbook = new Workbook(filePath, loadOptions);

                // Iterate through all worksheets and access their pivot tables
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    if (sheet.PivotTables.Count > 0)
                    {
                        Console.WriteLine($"Worksheet \"{sheet.Name}\" contains {sheet.PivotTables.Count} pivot table(s).");

                        // Process each pivot table in the worksheet
                        foreach (PivotTable pivotTable in sheet.PivotTables)
                        {
                            Console.WriteLine($"  Pivot Table Name: {pivotTable.Name}");

                            try
                            {
                                // Refresh the pivot cache data (correct API)
                                pivotTable.RefreshData();
                                // Recalculate the pivot table after refresh
                                pivotTable.CalculateData();
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"    Error refreshing pivot table \"{pivotTable.Name}\": {ex.Message}");
                            }

                            // Output the number of row fields
                            Console.WriteLine($"    Row Fields Count: {pivotTable.RowFields.Count}");
                        }
                    }
                }
            }
            catch (CellsException ex)
            {
                // Handles errors related to Aspose.Cells operations, including invalid password
                Console.WriteLine($"Aspose.Cells error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
