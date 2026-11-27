// Title: Batch apply a red holiday style to all date cells in multiple Excel workbooks using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells batch apply custom style to date cells in every worksheet of a folder | How to highlight calendar dates with a red background across many .xlsx files using Aspose.Cells | Apply holiday theme to date columns in bulk with Aspose.Cells .NET | Iterate through worksheets and set style for DateTime cells programmatically in C# | Save themed Excel workbooks to a separate output directory after styling date cells
// Tags: batch apply custom style to DateTime cells Aspose.Cells | iterate worksheets and cells to set style C# | save themed workbooks to output folder Aspose.Cells | holiday theme formatting for Excel .xlsx files | create solid fill cell style Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using System.Drawing;

// The C# program scans a folder for .xlsx workbooks, creates a red background with white bold font style, applies this style to every cell containing a DateTime value across all worksheets, and saves the updated workbooks to a designated output directory using Aspose.Cells for .NET.
class HolidayThemeBatchProcessor
{
    static void Main(string[] args)
    {
        // Input folder containing the workbooks to process
        string inputFolder = @"C:\Workbooks\Input";
        // Output folder where the themed workbooks will be saved
        string outputFolder = @"C:\Workbooks\Output";

        // Ensure output directory exists
        if (!Directory.Exists(outputFolder))
        {
            Directory.CreateDirectory(outputFolder);
        }

        // Process each .xlsx file in the input folder
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx"))
        {
            try
            {
                // Load the workbook (lifecycle rule: load)
                Workbook workbook = new Workbook(filePath);

                // Create a holiday style to apply to date cells
                Style holidayStyle = workbook.CreateStyle();
                holidayStyle.ForegroundColor = Color.Red;                     // Background color
                holidayStyle.Pattern = BackgroundType.Solid;                  // Solid fill
                holidayStyle.Font.Color = Color.White;                        // Font color
                holidayStyle.Font.IsBold = true;                              // Bold font
                holidayStyle.Font.Size = 11;                                  // Font size

                // Apply the style to all cells that contain dates in every worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;
                    // Determine the used range to limit iteration
                    int maxRow = cells.MaxDataRow;
                    int maxCol = cells.MaxDataColumn;

                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = cells[row, col];
                            // Check if the cell holds a DateTime value
                            if (cell.Type == CellValueType.IsDateTime)
                            {
                                // Apply the holiday style (chart-values rule not needed here)
                                cell.SetStyle(holidayStyle);
                            }
                        }
                    }
                }

                // Save the modified workbook to the output folder (lifecycle rule: save)
                string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
                workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("Holiday theme applied to all workbooks.");
    }
}
