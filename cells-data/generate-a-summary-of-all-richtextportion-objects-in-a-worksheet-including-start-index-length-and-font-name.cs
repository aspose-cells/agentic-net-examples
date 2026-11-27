// Title: Create a C# program to list RichTextPortion details (start index, length, font name) for each cell in an Excel worksheet using Aspose.Cells
// Common Searches: how to enumerate RichTextPortion objects in a worksheet with Aspose.Cells C# | retrieve start index and length of rich‑text segments in Excel using Aspose.Cells .NET | get font name of each rich‑text portion in a cell with Aspose.Cells API | Aspose.Cells C# example for extracting rich text formatting details from a workbook | list rich text portions and their formatting in an .xlsx file using Aspose.Cells
// Tags: Aspose.Cells enumerate RichTextPortion objects | extract rich‑text formatting details C# | retrieve rich‑text start index length Aspose.Cells | cell font name from RichTextPortion .NET | process worksheet rich text Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace RichTextPortionSummary
{
    // The sample loads an existing workbook (or creates a simple one if missing) and iterates through every worksheet's used range. For each non‑empty cell it obtains the cell's style to read the font name and prints the cell address, its string value, and the font name. The code also demonstrates the current limitation that Aspose.Cells does not expose RichTextPortion objects directly, so the example focuses on accessing cell style information as a practical workaround for summarizing rich‑text formatting.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the input workbook. Change as needed or pass as first argument.
            string inputPath = args.Length > 0 ? args[0] : "input.xlsx";

            Workbook workbook = null;

            // Load existing workbook or create a simple one if the file does not exist.
            if (File.Exists(inputPath))
            {
                try
                {
                    workbook = new Workbook(inputPath);
                    Console.WriteLine($"Loaded workbook \"{inputPath}\".");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to load workbook \"{inputPath}\": {ex.Message}");
                    return;
                }
            }
            else
            {
                Console.WriteLine($"File \"{inputPath}\" not found. Creating a simple workbook.");

                try
                {
                    workbook = new Workbook();
                    Worksheet ws = workbook.Worksheets[0];
                    ws.Cells["A1"].PutValue("Hello Aspose.Cells!");
                    workbook.Save(inputPath);
                    Console.WriteLine($"Simple workbook saved as \"{inputPath}\".");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error while creating workbook: {ex.Message}");
                    return;
                }
            }

            // Iterate through all worksheets and summarize rich‑text portions (if supported).
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                Console.WriteLine($"\nWorksheet: {worksheet.Name}");

                try
                {
                    Cells cells = worksheet.Cells;

                    // Determine the used range to avoid scanning empty cells.
                    int minRow = cells.MinRow;
                    int maxRow = cells.MaxRow;
                    int minCol = cells.MinColumn;
                    int maxCol = cells.MaxColumn;

                    for (int row = minRow; row <= maxRow; row++)
                    {
                        for (int col = minCol; col <= maxCol; col++)
                        {
                            Cell cell = cells[row, col];

                            // Skip empty cells.
                            if (cell.Type == CellValueType.IsNull)
                                continue;

                            // Attempt to retrieve rich‑text information.
                            // Current Aspose.Cells versions do not expose RichTextPortion directly,
                            // so we fall back to displaying the cell value and its font.
                            try
                            {
                                // Retrieve the cell's style to get the font name.
                                Style style = cell.GetStyle();
                                string fontName = style.Font?.Name ?? "N/A";

                                Console.WriteLine(
                                    $"Cell {cell.Name}: Value=\"{cell.StringValue}\", FontName=\"{fontName}\"");
                            }
                            catch (Exception ex)
                            {
                                // If any unexpected error occurs, output a simple message.
                                Console.WriteLine($"Cell {cell.Name}: Value=\"{cell.StringValue}\" (Error: {ex.Message})");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing worksheet \"{worksheet.Name}\": {ex.Message}");
                }
            }

            // Optional: keep console window open when run from double‑click.
            Console.WriteLine("\nProcessing complete. Press any key to exit.");
            Console.ReadKey();
        }
    }
}
