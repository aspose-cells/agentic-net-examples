// Title: Copy specific rows from an Excel worksheet to a new workbook using Aspose.Cells LightCells with low memory usage (C#)
// Common Searches: C# Aspose.Cells LightCells copy only certain rows to a new XLSX file | How to export selected rows from a large Excel workbook without loading the entire sheet in C# | Using LoadOptions MemoryPreference to copy rows while preserving formatting in Aspose.Cells | Minimal memory row copy example with Aspose.Cells LightCells API | Copy rows preserving cell styles to another workbook using Aspose.Cells C#
// Tags: LightCells copy rows to new workbook | MemoryPreference load options Aspose.Cells | preserve cell style during row copy C# | selective row extraction Excel Aspose.Cells | low‑memory Excel row transfer C#

using System;
using System.IO;
using Aspose.Cells;

namespace LightCellsCopyRows
{
    // The example loads a source XLSX workbook with LoadOptions set to MemoryPreference to minimize memory consumption, iterates over a predefined list of row indices, copies each cell's value and style to a new workbook using LightCells, and saves the result as selected_rows.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the source workbook
                string sourcePath = "source.xlsx";

                // Verify that the source file exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Define the zero‑based row indices that should be copied
                int[] rowsToCopy = new int[] { 1, 3, 5 }; // example rows

                // Load the source workbook with memory‑optimized settings
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    MemorySetting = MemorySetting.MemoryPreference // minimizes memory usage
                };
                Workbook sourceWorkbook = new Workbook(sourcePath, loadOptions);
                Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

                // Create a new workbook that will receive the selected rows
                Workbook destWorkbook = new Workbook();
                Worksheet destSheet = destWorkbook.Worksheets[0];

                // Determine the number of columns in the source sheet (rectangular range)
                int columnCount = sourceSheet.Cells.MaxColumn + 1;

                int destRowIndex = 0; // current row in the destination sheet

                // Iterate over the rows that need to be copied
                foreach (int srcRowIndex in rowsToCopy)
                {
                    for (int col = 0; col < columnCount; col++)
                    {
                        // Get source and destination cells
                        Cell srcCell = sourceSheet.Cells[srcRowIndex, col];
                        Cell destCell = destSheet.Cells[destRowIndex, col];

                        // Copy cell value
                        destCell.PutValue(srcCell.Value);

                        // Copy cell style (font, color, borders, etc.)
                        Style srcStyle = srcCell.GetStyle();
                        destCell.SetStyle(srcStyle);
                    }

                    // Move to the next row in the destination sheet
                    destRowIndex++;
                }

                // Save the new workbook containing only the selected rows
                destWorkbook.Save("selected_rows.xlsx");
                Console.WriteLine("Selected rows have been saved to 'selected_rows.xlsx'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
