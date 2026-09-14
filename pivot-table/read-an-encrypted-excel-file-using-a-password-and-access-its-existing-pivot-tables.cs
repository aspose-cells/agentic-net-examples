// Title: Read a password‑protected .xlsx file and list its pivot tables using Aspose.Cells for .NET
// AI Prompts: Open a password‑protected Excel file with Aspose.Cells LoadOptions and retrieve every PivotTable from each worksheet. | After loading the protected workbook, call RefreshData and CalculateData on each PivotTable and print the number of row and data fields. | Add robust error handling for missing files, wrong passwords, and pivot‑table processing failures when working with an encrypted Excel workbook in C#.
// Common Searches: Aspose.Cells C# read encrypted Excel workbook and get pivot tables | open password protected .xlsx with Aspose.Cells and list all pivot tables | C# Aspose.Cells LoadOptions password example for accessing pivot tables | refresh pivot table data after loading a protected Excel file using Aspose.Cells
// Tags: load workbook with password Aspose.Cells | enumerate pivot tables Aspose.Cells .NET | recalculate pivot tables Aspose.Cells | pivot field count extraction Aspose.Cells C# | encrypted Excel exception handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsEncryptedPivotDemo
{
    // Demonstrates opening a password‑protected Excel workbook via Aspose.Cells LoadOptions, iterating worksheets to access each PivotTable, refreshing and calculating its data, and outputting row and data field counts while handling possible errors.
    class Program
    {
        static void Main()
        {
            // Path to the encrypted Excel file
            string encryptedFilePath = "encrypted_with_pivot.xlsx";

            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(encryptedFilePath))
            {
                Console.WriteLine($"File not found: {encryptedFilePath}");
                return;
            }

            // Password used to protect the workbook
            string password = "mySecretPwd";

            try
            {
                // Load the workbook with the password using LoadOptions
                LoadOptions loadOptions = new LoadOptions
                {
                    Password = password
                };

                // Open the encrypted workbook
                Workbook workbook = new Workbook(encryptedFilePath, loadOptions);

                // Optional: verify that the workbook is indeed encrypted
                Console.WriteLine("Workbook is encrypted: " + workbook.Settings.IsEncrypted);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Access the collection of pivot tables in the current worksheet
                    PivotTableCollection pivots = sheet.PivotTables;

                    // If there are pivot tables, process them
                    if (pivots.Count > 0)
                    {
                        Console.WriteLine($"Worksheet \"{sheet.Name}\" contains {pivots.Count} pivot table(s).");

                        for (int i = 0; i < pivots.Count; i++)
                        {
                            try
                            {
                                PivotTable pivot = pivots[i];

                                // Refresh the pivot data and recalculate
                                pivot.RefreshData();
                                pivot.CalculateData();

                                // Output basic information about the pivot table
                                Console.WriteLine($"  Pivot Table \"{pivot.Name}\":");
                                // Source data string may not be directly exposed in newer API versions; skip if unavailable
                                Console.WriteLine($"    Row Fields Count: {pivot.RowFields.Count}");
                                Console.WriteLine($"    Data Fields Count: {pivot.DataFields.Count}");
                            }
                            catch (Exception exPivot)
                            {
                                Console.WriteLine($"    Error processing pivot table at index {i}: {exPivot.Message}");
                            }
                        }
                    }
                }

                // No need to save if only reading; if modifications are required, uncomment below:
                // workbook.Save("output.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
