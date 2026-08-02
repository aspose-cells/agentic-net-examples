// Title: Add a Calculated Column with IF Formula in Aspose.Cells for C# (.NET)
// Description: Creates a new workbook, populates columns A and B with sample items and numbers, defines a numeric threshold, inserts an =IF formula into column C to tag each row as “High” or “Low”, evaluates all formulas, and saves the file as WorkbookWithCalculatedColumn.xlsx.
// Keywords: Aspose.Cells | C# | .NET | IF formula | calculated column | Excel automation | threshold classification | programmatic formula | evaluate formulas | save workbook | Excel data categorization
// Common Searches: Aspose.Cells add calculated column with IF | C# set IF formula in Excel using Aspose.Cells | programmatically categorize rows by threshold in .xlsx | how to evaluate formulas after adding them in Aspose.Cells | save workbook after inserting calculated column Aspose.Cells
// Developer Intent: Programmatically create an Excel worksheet, add an IF‑based calculated column that classifies numeric values against a threshold, recalculate the sheet, and write the result to a file.
// Use Cases: Automatic scoring of KPI values as “High” or “Low” in a generated report. | Data‑validation sheet that flags entries exceeding a business limit without manual edits. | Export of processed analytics where each record includes a pre‑computed category for downstream tools.
// AI Prompts: Generate C# code using Aspose.Cells that adds a calculated column with an IF formula comparing each row's value to a given threshold and saves the workbook. | Show how to loop through rows in Aspose.Cells, assign an =IF formula to a cell, recalculate all formulas, and export the result as an .xlsx file. | Explain how to define a numeric threshold variable, build the IF formula with string interpolation, and apply it to multiple rows in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a new workbook, populates columns A and B with sample items and numbers, defines a numeric threshold, inserts an =IF formula into column C to tag each row as “High” or “Low”, evaluates all formulas, and saves the file as WorkbookWithCalculatedColumn.xlsx.
    public class AddCalculatedColumnWithIfDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Sample data: Column A = Item, Column B = Value
                cells["A1"].PutValue("Item");
                cells["B1"].PutValue("Value");
                cells["A2"].PutValue("Item1");
                cells["B2"].PutValue(80);
                cells["A3"].PutValue("Item2");
                cells["B3"].PutValue(120);
                cells["A4"].PutValue("Item3");
                cells["B4"].PutValue(95);
                cells["A5"].PutValue("Item4");
                cells["B5"].PutValue(150);

                // Header for the calculated column
                cells["C1"].PutValue("Category");

                // Define the threshold
                double threshold = 100;

                // Add IF formula to each data row to categorize based on the threshold
                for (int row = 2; row <= 5; row++)
                {
                    // Formula: =IF(B[row] > threshold, "High", "Low")
                    string formula = $"=IF(B{row}>{threshold},\"High\",\"Low\")";
                    cells[$"C{row}"].Formula = formula;
                }

                // Calculate all formulas so that the result values are stored
                workbook.CalculateFormula();

                // Save the workbook
                string outputPath = "WorkbookWithCalculatedColumn.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
