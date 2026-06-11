using System;
using Aspose.Cells;

namespace FreezePanesDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Expect two command‑line arguments: row index and column index (zero‑based)
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: FreezePanesDemo <rowIndex> <columnIndex>");
                return;
            }

            // Parse arguments
            if (!int.TryParse(args[0], out int rowIndex) || !int.TryParse(args[1], out int columnIndex))
            {
                Console.WriteLine("Both arguments must be valid integers.");
                return;
            }

            // Create a new workbook (empty)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Freeze panes at the specified cell.
            // The frozenRows and frozenColumns are set to the same values as the position,
            // which means rows 0..rowIndex‑1 and columns 0..columnIndex‑1 will be frozen.
            worksheet.FreezePanes(rowIndex, columnIndex, rowIndex, columnIndex);

            // Optional: indicate the freeze state
            worksheet.GetFreezedPanes(out int r, out int c, out int fr, out int fc);
            Console.WriteLine($"Freeze applied at row {r}, column {c} with {fr} frozen rows and {fc} frozen columns.");

            // Save the workbook
            string outputPath = "FreezePanesResult.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}