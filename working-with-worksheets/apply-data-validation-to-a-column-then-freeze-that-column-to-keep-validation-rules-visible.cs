// Title: Add list‑type data validation to column A and freeze that column with Aspose.Cells for .NET (C#)
// Common Searches: C# code to set list data validation for cells A1 to A100 using Aspose.Cells | How to freeze column A in an Excel workbook with Aspose.Cells .NET | Example of applying data validation and FreezePanes together in Aspose.Cells C# | Aspose.Cells create workbook with dropdown list validation and frozen first column | Set dropdown list validation and keep the column visible while scrolling in Excel via Aspose.Cells
// Tags: Aspose.Cells list validation C# | freeze first column Aspose.Cells | data validation range A1:A100 Aspose.Cells | worksheet FreezePanes method C# | Excel dropdown validation with frozen column Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDataValidationFreeze
{
    // The sample creates a new workbook, adds a list‑type data validation (Option1, Option2, Option3) to cells A1‑A100 on the first worksheet, freezes column A using the FreezePanes method, and saves the file as DataValidationFreeze.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the output Excel file
            string outputPath = "DataValidationFreeze.xlsx";

            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];
                worksheet.Name = "SampleData";

                // Define the range for the validation (A1:A100)
                CellArea validationArea = CellArea.CreateCellArea("A1", "A100");

                // Add a validation rule to the specified range
                int validationIndex = worksheet.Validations.Add(validationArea);
                Validation validation = worksheet.Validations[validationIndex];

                // Configure the validation as a list with three allowed values
                validation.Type = ValidationType.List;
                // The list values must be enclosed in double quotes and separated by commas
                validation.Formula1 = "\"Option1,Option2,Option3\"";

                // Freeze the first column (column A) so that it stays visible while scrolling
                // FreezePanes(row, column, topRow, leftColumn) – rows above 'row' and columns left of 'column' are frozen
                worksheet.FreezePanes(0, 1, 0, 0);

                // Ensure the directory for the output file exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to the specified file
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine("Data validation applied to column A and the column has been frozen.");
            }
            catch (Exception ex)
            {
                // Output any errors that occur during processing
                Console.WriteLine("An error occurred while creating the workbook:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
