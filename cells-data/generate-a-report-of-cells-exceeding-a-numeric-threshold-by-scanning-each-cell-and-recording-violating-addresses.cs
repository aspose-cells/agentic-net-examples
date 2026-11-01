// Title: Generate an Excel threshold violation report of numeric cells using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# how to list cells with values greater than a specific number | Create a summary worksheet for out-of-range numeric values in an Excel file using Aspose.Cells | C# code to scan all worksheets and report cells exceeding a threshold with Aspose.Cells | Generate a threshold breach report in Excel programmatically with Aspose.Cells .NET | Find and export cell addresses that exceed a numeric limit using Aspose.Cells in C#
// Tags: threshold violation report Aspose.Cells C# | scan numeric cells Aspose.Cells | create summary worksheet out-of-range values | iterate MaxDisplayRange Aspose.Cells | export violating cell addresses to new sheet

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace ThresholdReport
{
    // The example defines a numeric threshold, loads an existing workbook, iterates through each worksheet's used range via MaxDisplayRange, checks every cell for a numeric value exceeding the limit, records the full sheet!cell address, adds a new worksheet named "ThresholdReport", writes headers and the violating addresses with their values, and saves the workbook as output.xlsx while handling missing files and runtime errors.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Define the numeric threshold
                double threshold = 100.0;

                // Verify input file exists
                const string inputPath = "input.xlsx";
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // List to store addresses of cells that exceed the threshold
                List<string> violatingAddresses = new List<string>();

                // Iterate through each worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Get the used range of the worksheet (fully qualified to avoid ambiguity)
                    Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;

                    // Loop through each cell in the used range
                    foreach (Cell cell in usedRange)
                    {
                        // Check if the cell contains a numeric value
                        if (cell.Value != null && double.TryParse(cell.Value.ToString(), out double numericValue))
                        {
                            // If the numeric value exceeds the threshold, record its address
                            if (numericValue > threshold)
                            {
                                // cell.Name returns the address in A1 style (e.g., "B3")
                                violatingAddresses.Add($"{sheet.Name}!{cell.Name}");
                            }
                        }
                    }
                }

                // Create a new worksheet to hold the report
                int reportIndex = workbook.Worksheets.Add();
                Worksheet reportSheet = workbook.Worksheets[reportIndex];
                reportSheet.Name = "ThresholdReport";

                // Write header
                reportSheet.Cells[0, 0].PutValue("Cell Address");
                reportSheet.Cells[0, 1].PutValue("Value");

                // Populate the report with violating cell addresses and their values
                int rowIndex = 1;
                foreach (string address in violatingAddresses)
                {
                    // Split sheet name and cell address
                    string[] parts = address.Split('!');
                    if (parts.Length != 2) continue; // safety check

                    string sheetName = parts[0];
                    string cellAddress = parts[1];

                    // Retrieve the actual cell to get its value
                    Worksheet sourceSheet = workbook.Worksheets[sheetName];
                    Cell sourceCell = sourceSheet.Cells[cellAddress];

                    // Write address and value to the report sheet
                    reportSheet.Cells[rowIndex, 0].PutValue(address);
                    reportSheet.Cells[rowIndex, 1].PutValue(sourceCell.Value);
                    rowIndex++;
                }

                // Save the workbook with the report
                const string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Report generated and saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
