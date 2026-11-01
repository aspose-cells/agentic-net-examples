// Title: Import CSV data into an existing Excel workbook at cell D4 and save as XLSX using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells import CSV file into existing worksheet starting at D4 | How to load an existing XLSX and insert CSV rows at a specific cell with Aspose.Cells | Aspose.Cells example for reading CSV and writing to cell D4 then saving as XLSX | Insert CSV data into a particular range of an Excel file using Aspose.Cells C# | Save modified workbook after adding CSV content at D4 with Aspose.Cells for .NET
// Tags: import csv into existing workbook Aspose.Cells | write csv values starting at cell D4 | save workbook as xlsx Aspose.Cells | load existing worksheet and insert csv rows C# | csv to excel conversion using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The C# program loads an existing XLSX workbook, reads a CSV file line by line, writes each CSV field into the worksheet beginning at cell D4, optionally skips a header row, and then saves the updated workbook as a new XLSX file, handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string workbookPath = "ExistingWorkbook.xlsx";
            const string csvPath = "SourceData.csv";
            const string resultPath = "ResultWorkbook.xlsx";

            // Verify that required files exist
            if (!File.Exists(workbookPath))
                throw new FileNotFoundException($"Workbook file not found: {workbookPath}");
            if (!File.Exists(csvPath))
                throw new FileNotFoundException($"CSV file not found: {csvPath}");

            // Load the existing workbook
            Workbook workbook = new Workbook(workbookPath);

            // Get the target worksheet (first worksheet in this example)
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the start cell (D4 => row index 3, column index 3)
            int startRow = 3;
            int startColumn = 3;

            // Read all lines from the CSV file
            string[] csvLines = File.ReadAllLines(csvPath);

            // Optionally skip header row (uncomment if needed)
            // bool hasHeader = true;
            // int dataStartIndex = hasHeader ? 1 : 0;

            int dataStartIndex = 0; // change to 1 if the first row is a header

            // Import CSV data into the worksheet
            for (int i = dataStartIndex; i < csvLines.Length; i++)
            {
                string line = csvLines[i];
                // Split by comma (basic CSV parsing; adjust if your CSV uses different delimiters or quoting)
                string[] values = line.Split(',');

                for (int j = 0; j < values.Length; j++)
                {
                    // Write each value to the corresponding cell
                    worksheet.Cells[startRow + i - dataStartIndex, startColumn + j].PutValue(values[j]);
                }
            }

            // Save the workbook as XLSX
            workbook.Save(resultPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{resultPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
