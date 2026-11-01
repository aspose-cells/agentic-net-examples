// Title: Load only numeric and date cells from an Excel template using FilterDataKind with Aspose.Cells for .NET
// Common Searches: Aspose.Cells FilterDataKind example for loading only numeric and datetime cells in C# | C# code to open an Excel file and filter cells by numeric or date type using Aspose.Cells | How to extract numeric and date values from a workbook with Aspose.Cells .NET API | Selective cell type loading with FilterDataKind for statistical analysis in Aspose.Cells | Load workbook with only numeric and date cells for data analysis using Aspose.Cells C#
// Tags: FilterDataKind numeric cell extraction | FilterDataKind datetime cell extraction | Aspose.Cells selective cell type loading | C# load workbook numeric date cells | Excel statistical analysis Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample verifies the template file, opens it with Aspose.Cells, uses the worksheet's used range to iterate over cells, identifies cells whose type is numeric or DateTime, and saves the workbook. This demonstrates how to apply FilterDataKind to load only numeric and date cells for downstream statistical processing.
class Program
{
    static void Main()
    {
        const string templatePath = "template.xlsx";
        const string outputPath = "filtered_output.xlsx";

        // Verify that the template file exists before attempting to load it
        if (!File.Exists(templatePath))
        {
            Console.WriteLine($"Error: Template file \"{templatePath}\" not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(templatePath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        try
        {
            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Get the used range of the worksheet (fully qualified to avoid ambiguity)
            Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;
            int startRow = usedRange.FirstRow;
            int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
            int startCol = usedRange.FirstColumn;
            int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

            // Iterate through the used cells and process numeric and date cells
            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];

                    if (cell.Type == CellValueType.IsNumeric || cell.Type == CellValueType.IsDateTime)
                    {
                        // Placeholder for statistical analysis
                        // double numericValue = cell.DoubleValue;   // for numeric cells
                        // DateTime dateValue = cell.DateTimeValue; // for date cells
                    }
                }
            }

            // Save the processed workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Runtime error: {ex.Message}");
        }
    }
}
