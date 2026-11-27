// Title: C# tool using Aspose.Cells to highlight and map circular reference cells in an Excel workbook
// Common Searches: how to detect circular references with Aspose.Cells in C# | C# Aspose.Cells highlight cells that cause circular reference errors | create a diagnostic worksheet that visualizes circular reference chains using Aspose.Cells | extract cell addresses from CellsException message in Aspose.Cells C# code | add comments to Excel cells programmatically with Aspose.Cells | copy original sheet data and mark circular reference cells in a new worksheet
// Tags: Aspose.Cells detect circular references | Aspose.Cells highlight circular reference cells | Aspose.Cells copy worksheet data | Aspose.Cells add cell comments | Aspose.Cells extract addresses from exception | C# Excel circular reference visualization

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook with Aspose.Cells, runs CalculateFormula to trigger any circular reference exceptions, extracts the offending cell addresses from the exception message and from #REF! errors, creates a new worksheet, copies the original data, highlights each circular cell with a LightSalmon background, adds sequential comments, and saves the file as a diagnostic map of the circular reference chain.
class Program
{
    static void Main(string[] args)
    {
        // Input and output file handling
        string inputPath = args.Length > 0 ? args[0] : "input.xlsx";
        string outputPath = args.Length > 1 ? args[1] : "output_circular_map.xlsx";

        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" does not exist.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            List<string> circularCellAddresses = new List<string>();

            // Attempt to calculate all formulas and capture circular reference exceptions
            try
            {
                workbook.CalculateFormula();
            }
            catch (CellsException ex)
            {
                // Aspose.Cells may throw a CellsException with a message containing cell addresses
                Console.WriteLine("Calculation exception caught: " + ex.Message);
                circularCellAddresses.AddRange(ExtractCellAddresses(ex.Message));
            }

            // Scan all cells for #REF! errors which often indicate circular references
            foreach (Worksheet ws in workbook.Worksheets)
            {
                Cells cells = ws.Cells;
                foreach (Cell cell in cells)
                {
                    if (cell.Type == CellValueType.IsError && cell.StringValue == "#REF!")
                    {
                        string address = cell.Name;
                        if (!circularCellAddresses.Contains(address))
                        {
                            circularCellAddresses.Add(address);
                        }
                    }
                }
            }

            if (circularCellAddresses.Count == 0)
            {
                Console.WriteLine("No circular references detected.");
                // Still save a copy of the original workbook for consistency
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
                return;
            }

            // Create a new worksheet to visualize the circular reference chain
            Worksheet mapSheet = workbook.Worksheets[workbook.Worksheets.Add()];
            mapSheet.Name = "CircularReferenceMap";

            // Copy original data to the map sheet (optional: you can copy only the involved cells)
            // Here we copy the entire first worksheet for context
            if (workbook.Worksheets.Count > 1)
            {
                Worksheet source = workbook.Worksheets[0];
                mapSheet.Cells.CopyRows(source.Cells, 0, source.Cells.MaxDataRow + 1, 0);
                mapSheet.Cells.CopyColumns(source.Cells, 0, source.Cells.MaxDataColumn + 1, 0);
            }

            // Apply visual markers to each circular cell
            int index = 1;
            foreach (string address in circularCellAddresses)
            {
                // Resolve the cell in the map sheet (same address as original)
                Cell mapCell = mapSheet.Cells[address];

                // Preserve the original value (if any)
                // (Already copied above; if not copied, you could fetch from original sheet)

                // Apply background color
                Style style = mapCell.GetStyle();
                style.ForegroundColor = System.Drawing.Color.LightSalmon;
                style.Pattern = BackgroundType.Solid;
                mapCell.SetStyle(style);

                // Add a comment indicating the order in the chain
                int row = mapCell.Row;
                int column = mapCell.Column;
                int commentIndex = mapSheet.Comments.Add(row, column);
                Comment comment = mapSheet.Comments[commentIndex];
                comment.Note = $"Circular reference #{index}";
                comment.Width = 200;
                comment.Height = 50;

                Console.WriteLine($"Marked circular cell {address} as #{index}");
                index++;
            }

            // Save the workbook with visualization
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Circular reference map saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Unexpected error: " + ex.Message);
            Console.WriteLine(ex.StackTrace);
        }
    }

    private static IEnumerable<string> ExtractCellAddresses(string text)
    {
        var addresses = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (string.IsNullOrEmpty(text))
            return addresses;

        // Simple pattern: one or more letters followed by one or more digits
        Regex regex = new Regex(@"\b[A-Z]{1,3}\d{1,7}\b", RegexOptions.IgnoreCase);
        foreach (Match match in regex.Matches(text))
        {
            addresses.Add(match.Value.ToUpperInvariant());
        }
        return addresses;
    }
}
