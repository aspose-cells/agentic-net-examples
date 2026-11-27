// Title: C# Example: Sorting Excel Worksheet Rows by Column A (Ascending) Using Aspose.Cells DataSorter
// Common Searches: how to sort an Excel worksheet by the first column using Aspose.Cells in C# | Aspose.Cells DataSorter sort rows ascending column A example | C# code to sort rows in an .xlsx file based on column A with Aspose | using CellArea with Aspose.Cells DataSorter to sort worksheet data | load or create workbook then sort rows by column A Aspose.Cells C#
// Tags: Aspose.Cells DataSorter ascending column sort | C# sort Excel rows by first column | define CellArea range for Aspose.Cells sorting | save sorted workbook Aspose.Cells C# | initialize workbook with sample data Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing workbook or creates a new one, determines the used range, configures Aspose.Cells.DataSorter to sort rows by column A in ascending order, applies the sort to a CellArea covering the data, saves the sorted file as 'sorted.xlsx', and prints the sorted rows to the console.
class Program
{
    static void Main()
    {
        // Paths for input and output files
        string inputPath = "input.xlsx";
        string outputPath = "sorted.xlsx";

        Workbook workbook;

        // Load existing workbook if it exists; otherwise create a sample workbook
        if (File.Exists(inputPath))
        {
            try
            {
                workbook = new Workbook(inputPath);
                Console.WriteLine($"Loaded workbook from '{inputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load '{inputPath}': {ex.Message}");
                return;
            }
        }
        else
        {
            // Create a new workbook with sample data
            workbook = new Workbook();
            Worksheet ws = workbook.Worksheets[0];

            ws.Cells["A1"].PutValue("Name");
            ws.Cells["B1"].PutValue("Age");

            ws.Cells["A2"].PutValue("John");
            ws.Cells["B2"].PutValue(25);

            ws.Cells["A3"].PutValue("Alice");
            ws.Cells["B3"].PutValue(30);

            ws.Cells["A4"].PutValue("Bob");
            ws.Cells["B4"].PutValue(22);

            // Save the sample workbook for reference
            try
            {
                workbook.Save(inputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Sample workbook created and saved as '{inputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save sample workbook: {ex.Message}");
                return;
            }
        }

        Worksheet worksheet = workbook.Worksheets[0];

        // Determine the last row that contains data (zero‑based index)
        int lastDataRow = worksheet.Cells.MaxDataRow;
        if (lastDataRow < 0)
        {
            Console.WriteLine("The worksheet contains no data to sort.");
            return;
        }

        // Define the range to be sorted (including header row)
        CellArea sortArea = new CellArea
        {
            StartRow = 0,
            EndRow = lastDataRow,
            StartColumn = 0,
            EndColumn = worksheet.Cells.MaxDataColumn // include all used columns
        };

        // Configure the DataSorter: sort by column A (index 0) in ascending order
        workbook.DataSorter.Key1 = 0;               // Column A
        workbook.DataSorter.Order1 = SortOrder.Ascending;

        try
        {
            // Perform the sort
            workbook.DataSorter.Sort(worksheet.Cells, sortArea);

            // Save the sorted workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Rows sorted by column A and saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during sorting: {ex.Message}");
            return;
        }

        // Output the sorted data to the console
        Console.WriteLine("Sorted rows (including header):");
        for (int row = sortArea.StartRow; row <= sortArea.EndRow; row++)
        {
            string name = worksheet.Cells[row, 0].StringValue;
            string age = worksheet.Cells[row, 1].StringValue;
            Console.WriteLine($"{row}: {name}, {age}");
        }
    }
}
