// Title: How to count non‑empty cells in each worksheet row using Aspose.Cells Rows enumerator in C#
// AI Prompts: Write C# code with Aspose.Cells that loops through the Rows collection of a worksheet and prints the number of cells that contain a value for each row. | Modify the example to skip rows that have zero non‑empty cells and only display counts for rows that contain data. | Add functionality to store the per‑row non‑empty cell counts in a Dictionary<int,int> and write the results to a CSV file using Aspose.Cells.
// Common Searches: Aspose.Cells C# count filled cells per row in a worksheet | how to iterate rows and get non empty cell count with Aspose.Cells for .NET | skip empty rows when counting cells using Aspose.Cells Rows enumerator | C# Aspose.Cells example to output non‑null cell count per row | save workbook after analyzing row data with Aspose.Cells
// Tags: Aspose.Cells row-wise non‑empty cell tally | Rows collection enumeration in Aspose.Cells C# | worksheet row cell count extraction Aspose.Cells | CSV export of per‑row cell statistics Aspose.Cells | post‑analysis workbook saving Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a workbook, adds data to several cells, then iterates each worksheet row via the Rows collection. For each row it counts cells whose Type is not IsNull, prints the count, ensures the output directory exists, saves the workbook as Output.xlsx, and handles any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one if a template file is provided)
            Workbook workbook = new Workbook();

            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue(123);
            sheet.Cells["D2"].PutValue(DateTime.Now);
            sheet.Cells["A3"].PutValue("Test");
            sheet.Cells["C3"].PutValue(3.14);
            sheet.Cells["E3"].PutValue(true);

            // Enumerate rows and count non‑empty cells in each row
            foreach (Row row in sheet.Cells.Rows)
            {
                int nonEmptyCount = 0;

                // Row implements IEnumerable<Cell>; count cells that contain data
                foreach (Cell cell in row)
                {
                    if (cell.Type != CellValueType.IsNull)
                        nonEmptyCount++;
                }

                Console.WriteLine($"Row {row.Index + 1}: {nonEmptyCount} non‑empty cell(s)");
            }

            // Save the workbook
            string outputPath = "Output.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
