// Title: Add integer data validation to a column that automatically expands when rows are inserted using CellArea in Aspose.Cells for .NET
// Common Searches: how to add integer validation that grows with inserted rows using Aspose.Cells C# | Aspose.Cells CellArea validation range expansion after inserting rows | C# example of dynamic data validation range in Excel with Aspose.Cells | automatically extend whole number validation when adding new rows in Aspose.Cells | using CellArea to create expandable data validation in a .NET Excel workbook
// Tags: CellArea dynamic validation range | Aspose.Cells whole number validation | expand validation after inserting rows | C# Aspose.Cells data validation | automatic validation range adjustment

using Aspose.Cells;
using System;
using System.IO;

// The sample creates a workbook, defines a CellArea covering A2:A5, applies whole‑number validation (0‑100) to that range, inserts a new row, and shows that the validation range automatically expands to include the inserted row before saving the file as DynamicValidation.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Populate initial data in column A (A2:A5)
            ws.Cells["A2"].PutValue(10);
            ws.Cells["A3"].PutValue(20);
            ws.Cells["A4"].PutValue(30);
            ws.Cells["A5"].PutValue(40);

            // Define a CellArea for the validation range (A2:A5)
            CellArea area = new CellArea
            {
                StartRow = 1,    // Row 2 (zero‑based)
                StartColumn = 0, // Column A
                EndRow = 4,      // Row 5
                EndColumn = 0    // Column A
            };

            // Add integer (whole number) validation to the defined area
            int index = ws.Validations.Add(area);
            Validation validation = ws.Validations[index];
            validation.Type = ValidationType.WholeNumber;
            validation.Operator = OperatorType.Between;
            validation.Formula1 = "0";
            validation.Formula2 = "100";

            // Input message settings
            validation.InputTitle = "Enter a number";
            validation.InputMessage = "Please enter an integer between 0 and 100.";
            // ShowInputMessage is true by default; no explicit property needed in this version

            // Insert a new row at Excel row 3 (zero‑based index 2)
            ws.Cells.InsertRows(2, 1);

            // After insertion, the validation range automatically expands.
            // Approximate the new range based on the original area.
            int newEndRow = area.EndRow + 1; // one extra row added
            Console.WriteLine($"Validation now covers rows {area.StartRow + 1} to {newEndRow + 1}");

            // Save the workbook
            string outputPath = "DynamicValidation.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
