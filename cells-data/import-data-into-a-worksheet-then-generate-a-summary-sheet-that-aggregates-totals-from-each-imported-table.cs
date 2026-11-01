// Title: Import several CSV files into separate worksheets and create a summary sheet that totals column B using Aspose.Cells for .NET
// AI Prompts: Generate C# code that reads a list of CSV file paths, adds each as a new worksheet in an Aspose.Cells workbook, and writes a "Summary" worksheet that lists each sheet name with the sum of numeric values from column B. | Adjust the aggregation loop to compute the average of column B instead of the sum and update the summary worksheet to display the averages.
// Common Searches: how to load multiple CSV files into separate worksheets with Aspose.Cells C# | Aspose.Cells create summary worksheet that shows totals per sheet | sum values from column B across worksheets using Aspose.Cells .NET | remove default Sheet1 when adding new worksheets in Aspose.Cells | auto fit columns after writing data in Aspose.Cells workbook
// Tags: import csv files as individual worksheets Aspose.Cells | calculate per‑sheet column B sum Aspose.Cells | add summary sheet with aggregated results .NET | delete default empty worksheet Aspose.Cells | auto‑fit columns after data write Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsAggregation
{
    // The program creates a new workbook, imports each specified CSV file into its own worksheet, removes the default empty sheet, adds a "Summary" worksheet, and writes the sum of numeric values from column B of each imported sheet into the summary.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // ---------- Create a new workbook ----------
                var workbook = new Workbook(); // create workbook (rule)

                // List of CSV files to import
                var csvFiles = new List<string>
                {
                    "Data1.csv",
                    "Data2.csv",
                    "Data3.csv"
                };

                // Keep track of imported sheet names
                var importedSheetNames = new List<string>();

                // Import each CSV file into its own worksheet
                foreach (var csvFile in csvFiles)
                {
                    // Verify that the CSV file exists
                    if (!File.Exists(csvFile))
                    {
                        Console.WriteLine($"Warning: File \"{csvFile}\" not found. Skipping.");
                        continue;
                    }

                    try
                    {
                        // Load CSV into a temporary workbook
                        var loadOptions = new LoadOptions(LoadFormat.Csv);
                        var tempWb = new Workbook(csvFile, loadOptions); // load (rule)

                        // Get the first worksheet from the temporary workbook
                        var sourceSheet = tempWb.Worksheets[0];

                        // Rename the worksheet to match the source file (without extension)
                        var sheetName = Path.GetFileNameWithoutExtension(csvFile);
                        sourceSheet.Name = sheetName;

                        // Copy the worksheet into the main workbook using its name
                        workbook.Worksheets.AddCopy(sourceSheet.Name);
                        importedSheetNames.Add(sheetName);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing \"{csvFile}\": {ex.Message}");
                    }
                }

                // Remove the default empty sheet created with the workbook (if it exists)
                if (workbook.Worksheets.Count > 0 &&
                    workbook.Worksheets[0].Cells.MaxDataRow == -1 &&
                    workbook.Worksheets[0].Name == "Sheet1")
                {
                    workbook.Worksheets.RemoveAt(0);
                }

                // ---------- Create Summary Sheet ----------
                var summarySheet = workbook.Worksheets.Add("Summary");

                // Write headers
                summarySheet.Cells["A1"].PutValue("Table");
                summarySheet.Cells["B1"].PutValue("Total");

                int summaryRow = 1; // zero‑based index (row 2 in Excel)

                // Aggregate totals from each imported sheet
                foreach (var sheetName in importedSheetNames)
                {
                    var sheet = workbook.Worksheets[sheetName];

                    // Assume the numeric column is the second column (B)
                    int numericColumnIndex = 1; // zero‑based index for column B

                    // Find the last row with data in the sheet
                    int lastDataRow = sheet.Cells.MaxDataRow;

                    double total = 0.0;

                    // Sum values from row 2 (skip header) to the last data row
                    for (int row = 1; row <= lastDataRow; row++)
                    {
                        var cell = sheet.Cells[row, numericColumnIndex];
                        if (cell.Type == CellValueType.IsNumeric)
                        {
                            total += cell.DoubleValue;
                        }
                    }

                    // Write the table name and its total into the summary sheet
                    summarySheet.Cells[summaryRow, 0].PutValue(sheetName); // Column A
                    summarySheet.Cells[summaryRow, 1].PutValue(total);    // Column B

                    summaryRow++;
                }

                // Auto‑fit columns for better readability
                summarySheet.AutoFitColumns();

                // ---------- Save the workbook ----------
                string outputPath = "AggregatedReport.xlsx";
                workbook.Save(outputPath); // save (rule)

                Console.WriteLine($"Aggregation complete. File saved as {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
