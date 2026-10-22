// Title: How to batch‑add slicers to every pivot table in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that iterates over all worksheets and adds a slicer to each pivot table using Aspose.Cells, selecting the first available row, column, page, or data field. | Show how to programmatically place a slicer at a specific cell location for every pivot table in a workbook with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# loop through worksheets to insert slicers for all pivot tables | Automatically create slicer for first available pivot field using Aspose.Cells .NET | Add slicer to pivot tables when input Excel file may be missing Aspose.Cells
// Tags: batch add slicers Aspose.Cells .NET | pivot table slicer insertion C# | first available pivot field slicer Aspose.Cells | slicer placement at cell coordinates Excel .NET | fallback workbook creation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The program loads an existing workbook (or creates a new one if the file is absent), iterates through each worksheet and its pivot tables, picks the first available row, column, page, or data field, adds a slicer at the top‑left corner for that field, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load workbook if the input file exists; otherwise create a new empty workbook.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
            }

            // Iterate through all worksheets.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all pivot tables on the current worksheet.
                foreach (PivotTable pivot in sheet.PivotTables)
                {
                    // Choose a field name for the slicer: row, column, page, or data field.
                    string fieldName = null;
                    if (pivot.RowFields.Count > 0)
                        fieldName = pivot.RowFields[0].Name;
                    else if (pivot.ColumnFields.Count > 0)
                        fieldName = pivot.ColumnFields[0].Name;
                    else if (pivot.PageFields.Count > 0)
                        fieldName = pivot.PageFields[0].Name;
                    else if (pivot.DataFields.Count > 0)
                        fieldName = pivot.DataFields[0].Name;

                    // Add slicer if a suitable field was found.
                    if (!string.IsNullOrEmpty(fieldName))
                    {
                        // Add slicer at the top‑left corner (row 0, column 0).
                        sheet.Slicers.Add(pivot, 0, 0, fieldName);
                    }
                }
            }

            // Save the modified workbook.
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
