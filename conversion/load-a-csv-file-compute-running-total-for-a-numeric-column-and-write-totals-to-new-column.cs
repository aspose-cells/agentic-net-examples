// Title: Calculate a cumulative running total for a numeric column in a CSV file using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells read CSV and add cumulative total column | How to compute running sum of a column and export to new CSV with Aspose.Cells | Aspose.Cells .NET example for cumulative sum in CSV worksheet | Add running total column to CSV using Aspose.Cells workbook in C# | Save modified CSV with additional column using Aspose.Cells for .NET
// Tags: Aspose.Cells cumulative sum CSV | add running total column Aspose.Cells | CSV worksheet column calculation .NET | write new column to CSV Aspose.Cells | save modified CSV workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample checks for the input CSV, loads it into an Aspose.Cells Workbook, reads numeric values from a specified column, computes a running total, writes each cumulative sum into a new column with a header, and saves the result as a new CSV file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.csv";
            const string outputPath = "output.csv";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the CSV file into a workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (the CSV data is placed in the first sheet)
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Define the column indexes (0‑based)
            int numericColumnIndex = 1;   // e.g., column B contains the numbers
            int totalColumnIndex = 2;     // column C will hold the running totals

            // Write header for the totals column (optional)
            cells[0, totalColumnIndex].PutValue("Running Total");

            double runningTotal = 0.0;

            // Start from row 1 assuming row 0 is the header
            for (int row = 1; row <= cells.MaxDataRow; row++)
            {
                // Get the numeric value from the source column
                double value = cells[row, numericColumnIndex].DoubleValue;

                // Update the running total
                runningTotal += value;

                // Write the running total to the new column
                cells[row, totalColumnIndex].PutValue(runningTotal);
            }

            // Save the workbook with the new column as CSV
            workbook.Save(outputPath);
            Console.WriteLine($"Processing completed. Output saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
