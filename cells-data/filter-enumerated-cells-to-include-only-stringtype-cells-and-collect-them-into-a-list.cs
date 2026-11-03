// Title: How to filter and collect only string‑type cells from an Excel worksheet into a List<string> using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# iterate worksheet cells and get only string values | filter Excel cells by CellValueType.IsString and return list of strings | collect string cells from used range in .NET with Aspose.Cells | retrieve all text cells from a workbook using Aspose.Cells API | C# code to extract only textual cell values from an .xlsx file
// Tags: string type cell filtering Aspose.Cells | extract text values from worksheet Aspose.Cells | enumerate used range cells Aspose.Cells | collect Excel string cells into List<string> | C# workbook text cell extraction Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// Loads a workbook, iterates over the used range of the first worksheet, checks each cell's Type for string values, and adds the cell's StringValue to a List<string> that is returned.
public class StringCellCollector
{
    // Loads the workbook, enumerates all cells, filters string‑type cells and returns their values.
    public static List<string> CollectStringCells(string workbookPath)
    {
        // Prevent FileNotFoundException
        if (!File.Exists(workbookPath))
            throw new FileNotFoundException($"Workbook file not found: {workbookPath}");

        // Load the workbook (lifecycle rule)
        Workbook workbook = new Workbook(workbookPath);

        // Assume we work with the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        List<string> stringValues = new List<string>();

        // Iterate through the used range of cells
        int maxRow = cells.MaxDataRow;
        int maxColumn = cells.MaxDataColumn;

        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxColumn; col++)
            {
                Cell cell = cells[row, col];

                // Filter only cells that contain string values
                if (cell.Type == CellValueType.IsString)
                {
                    stringValues.Add(cell.StringValue);
                }
            }
        }

        return stringValues;
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            // Use first argument as workbook path or fallback to a default file name
            string workbookPath = args.Length > 0 ? args[0] : "sample.xlsx";

            List<string> strings = StringCellCollector.CollectStringCells(workbookPath);

            Console.WriteLine($"Collected {strings.Count} string cells:");
            foreach (var s in strings)
            {
                Console.WriteLine(s);
            }
        }
        catch (FileNotFoundException ex)
        {
            Console.Error.WriteLine(ex.Message);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
