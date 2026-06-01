using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;   // For ListObject
using AsposeRange = Aspose.Cells.Range; // Alias to avoid conflict with System.Range

namespace ExportTablesToCsv
{
    class Program
    {
        static void Main()
        {
            // Path to the source workbook containing tables
            string sourcePath = "input.xlsx";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            try
            {
                // Load the workbook inside a using block for deterministic disposal
                using (Workbook sourceWorkbook = new Workbook(sourcePath))
                {
                    // Iterate through each worksheet
                    foreach (Worksheet sheet in sourceWorkbook.Worksheets)
                    {
                        // Iterate through each table (ListObject) in the worksheet
                        foreach (ListObject table in sheet.ListObjects)
                        {
                            try
                            {
                                // Create a temporary workbook to hold the single table
                                using (Workbook tempWorkbook = new Workbook())
                                {
                                    // Get the first (and only) worksheet of the temporary workbook
                                    Worksheet tempSheet = tempWorkbook.Worksheets[0];

                                    // Determine the size of the table's data range (including header)
                                    int rows = table.DataRange.RowCount;
                                    int cols = table.DataRange.ColumnCount;

                                    // Create a destination range in the temporary sheet starting at A1
                                    AsposeRange destRange = tempSheet.Cells.CreateRange(0, 0, rows, cols);

                                    // Copy the table's data range to the destination range
                                    table.DataRange.Copy(destRange);

                                    // Prepare CSV save options – export only the active sheet
                                    TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
                                    {
                                        ExportAllSheets = false
                                    };

                                    // Use the table's display name for the output file (fallback to a GUID if empty)
                                    string tableName = !string.IsNullOrEmpty(table.DisplayName) ? table.DisplayName : $"Table_{Guid.NewGuid():N}";
                                    string outputFileName = $"{tableName}.csv";

                                    // Save the temporary workbook as a CSV file
                                    tempWorkbook.Save(outputFileName, csvOptions);
                                    Console.WriteLine($"Exported table '{tableName}' to '{outputFileName}'.");
                                }
                            }
                            catch (Exception exTable)
                            {
                                Console.WriteLine($"Failed to export table '{table.DisplayName}': {exTable.Message}");
                            }
                        }
                    }
                }

                Console.WriteLine("All tables have been processed.");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}