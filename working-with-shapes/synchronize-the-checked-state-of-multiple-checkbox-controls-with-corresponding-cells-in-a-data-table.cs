// Title: Add a linked CheckBox shape to each data row and bind its checked state to a Boolean column using Aspose.Cells for .NET
// AI Prompts: Write C# code that iterates over all worksheet rows, reads the Boolean value from column B, creates a 20 × 20 CheckBox shape in column C, sets the CheckBox.LinkedCell to the same B cell, and saves the workbook. | Generate a C# example that programmatically inserts a CheckBox shape per data row, links each CheckBox to its corresponding Boolean cell, ensures the initial checked state matches the cell value, and writes the result to a new Excel file with Aspose.Cells.
// Common Searches: Aspose.Cells C# add checkbox shape to each data row and link to cell | loop to insert linked checkboxes in Excel using Aspose.Cells .NET | synchronize Excel checkbox state with a Boolean column via Aspose.Cells | set LinkedCell property for CheckBox shapes in Aspose.Cells C#
// Tags: checkbox shape insertion Aspose.Cells | checkbox linked cell property Aspose.Cells | bind checkbox state to boolean column Aspose.Cells | row-wise checkbox creation Aspose.Cells | save workbook with linked checkboxes Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an existing workbook, walks through each data row starting at row 2, reads the true/false value from column B, adds a 20 × 20 CheckBox shape in column C, links the shape to the same B cell so the checkbox reflects the cell's Boolean value, and saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Data starts from the second row (zero‑based index 1)
            int dataStartRow = 1;
            int lastDataRow = sheet.Cells.MaxDataRow;

            // Add a linked CheckBox in column C for each data row
            for (int row = dataStartRow; row <= lastDataRow; row++)
            {
                try
                {
                    // Read the boolean value from column B
                    bool isChecked = sheet.Cells[row, 1].BoolValue;

                    // Add a CheckBox shape anchored to column C (index 2)
                    CheckBox checkBox = sheet.Shapes.AddCheckBox(row, 2, row, 2, 0, 0);

                    // Link the CheckBox to column B of the same row (e.g., "B2")
                    checkBox.LinkedCell = sheet.Cells[row, 1].Name;

                    // Set the initial checked state of the linked cell
                    sheet.Cells[row, 1].PutValue(isChecked);

                    // Appearance adjustments
                    checkBox.Width = 20;
                    checkBox.Height = 20;
                }
                catch (Exception exRow)
                {
                    Console.WriteLine($"Warning: Could not add CheckBox to row {row + 1}. {exRow.Message}");
                }
            }

            // Ensure output directory exists (if a directory part is present)
            string outputDir = Path.GetDirectoryName(outputPath) ?? string.Empty;
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
