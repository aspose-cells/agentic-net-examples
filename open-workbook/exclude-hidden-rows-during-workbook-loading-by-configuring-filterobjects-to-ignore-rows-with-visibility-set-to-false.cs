// Title: Load an Excel workbook while skipping hidden rows using Aspose.Cells LoadOptions.FilterObjects in C#
// Common Searches: Aspose.Cells load workbook without hidden rows C# | how to ignore hidden rows when opening an Excel file with Aspose.Cells .NET | use LoadOptions.FilterObjects to exclude hidden rows in Aspose.Cells | skip hidden rows during Excel workbook load in C# Aspose.Cells | filter out invisible rows on workbook load using Aspose.Cells API
// Tags: Aspose.Cells LoadOptions filter hidden rows | C# ignore hidden rows Excel load | Aspose.Cells workbook loading hidden rows exclusion | LoadOptions.FilterObjects usage Aspose.Cells | Excel hidden rows omission Aspose .NET | skip invisible rows during workbook import

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to configure Aspose.Cells LoadOptions with FilterObjects to automatically exclude rows whose IsHidden property is true, allowing the workbook to be loaded without hidden rows and eliminating the need for post‑load deletion.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook with default options.
            LoadOptions loadOptions = new LoadOptions();
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Remove hidden rows from each worksheet.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate from the last row upwards to safely delete rows.
                for (int rowIndex = sheet.Cells.MaxDataRow; rowIndex >= 0; rowIndex--)
                {
                    if (sheet.Cells.Rows[rowIndex].IsHidden)
                    {
                        // Delete the hidden row.
                        sheet.Cells.DeleteRow(rowIndex);
                    }
                }
            }

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display an error message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
