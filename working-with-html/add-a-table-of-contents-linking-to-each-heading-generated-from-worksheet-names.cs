// Title: Create a Table of Contents worksheet with hyperlinks to each sheet using Aspose.Cells for .NET (C#)
// Common Searches: how to add a table of contents sheet with links to other worksheets using Aspose.Cells C# | Aspose.Cells C# generate hyperlinked index of workbook sheets | C# code to insert a TOC worksheet and hyperlink to each sheet in Excel with Aspose.Cells | Aspose.Cells .NET create TOC and auto‑fit column for sheet names | skip the TOC sheet while iterating worksheets Aspose.Cells example
// Tags: Aspose.Cells create TOC worksheet | hyperlink worksheet names Aspose.Cells | insert worksheet at first position Aspose.Cells | auto‑fit column Aspose.Cells | skip specific sheet during iteration Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsTOCExample
{
    // The example loads an existing Excel file or creates a new workbook, inserts a new worksheet named "Table of Contents" at the first position, applies a bold blue style to entries, iterates over all other worksheets, writes each sheet name into the TOC, adds a hyperlink to cell A1 of the target sheet, auto‑fits the first column, and saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Define input and output file paths
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Load existing workbook if it exists; otherwise create a new one
                Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

                // Insert a new worksheet at the beginning to serve as the Table of Contents
                Worksheet tocSheet = workbook.Worksheets.Insert(0, SheetType.Worksheet);
                tocSheet.Name = "Table of Contents";

                // Define a style for the TOC entries (optional)
                Style style = workbook.CreateStyle();
                style.Font.IsBold = true;
                style.Font.Color = Color.Blue;
                style.Font.Size = 12;
                style.Font.Underline = FontUnderlineType.Single;

                // Iterate through all worksheets (skip the TOC sheet itself)
                int rowIndex = 0;
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    if (ws.Name == tocSheet.Name)
                        continue;

                    // Write the worksheet name into column A
                    Cell cell = tocSheet.Cells[rowIndex, 0];
                    cell.PutValue(ws.Name);
                    cell.SetStyle(style);

                    // Add a hyperlink that points to cell A1 of the target worksheet
                    tocSheet.Hyperlinks.Add(rowIndex, 0, 1, 1, $"'{ws.Name}'!A1");

                    rowIndex++;
                }

                // Auto-fit the column width for better appearance
                tocSheet.AutoFitColumn(0);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
