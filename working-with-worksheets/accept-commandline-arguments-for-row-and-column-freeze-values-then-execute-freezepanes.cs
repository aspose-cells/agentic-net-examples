// Title: Freeze specific rows and columns in an Excel workbook using Aspose.Cells for .NET via command‑line arguments (C#)
// Common Searches: c# console app Aspose.Cells freeze panes using row and column parameters from command line | how to set FreezePanes on the first worksheet with dynamic indices in Aspose.Cells .NET | run Aspose.Cells program to freeze top rows and left columns based on user‑provided arguments
// Tags: Aspose.Cells FreezePanes command line | C# load workbook set freeze panes | Excel freeze rows columns Aspose.Cells | console application dynamic freeze indices | first worksheet FreezePanes method Aspose.Cells

using System;
using Aspose.Cells;

// A C# console application that reads input and output Excel file paths plus row and column indices from command‑line arguments, loads the workbook with Aspose.Cells, applies FreezePanes on the first worksheet using the supplied indices, and saves the modified file.
class Program
{
    static void Main(string[] args)
    {
        // Expected arguments:
        // args[0] - input Excel file path
        // args[1] - output Excel file path
        // args[2] - row index to freeze (0‑based)
        // args[3] - column index to freeze (0‑based)
        if (args.Length < 4)
        {
            Console.WriteLine("Usage: program.exe <input.xlsx> <output.xlsx> <freezeRow> <freezeColumn>");
            return;
        }

        string inputPath = args[0];
        string outputPath = args[1];
        int freezeRow = int.Parse(args[2]);
        int freezeColumn = int.Parse(args[3]);

        // Create workbook using the provided rule
        Workbook workbook = CreateWorkbook();

        // Load workbook using the provided rule
        workbook = LoadWorkbook(inputPath);

        // Apply FreezePanes on the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        // Freeze rows above freezeRow and columns left of freezeColumn
        // Parameters: row, column, topRow, leftColumn
        sheet.FreezePanes(freezeRow, freezeColumn, freezeRow, freezeColumn);

        // Save workbook using the provided rule
        SaveWorkbook(workbook, outputPath);
    }

    // ---- Lifecycle rule: create workbook ----
    static Workbook CreateWorkbook()
    {
        // {CreateWorkbook}
        return new Workbook();
    }

    // ---- Lifecycle rule: load workbook ----
    static Workbook LoadWorkbook(string path)
    {
        // {LoadWorkbook}
        return new Workbook(path);
    }

    // ---- Lifecycle rule: save workbook ----
    static void SaveWorkbook(Workbook wb, string path)
    {
        // {SaveWorkbook}
        wb.Save(path);
    }
}
