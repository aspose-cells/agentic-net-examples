// Title: How to merge cells D2:E2 and add a dropdown list validation in an Excel file using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells merge D2:E2 and create a dropdown list validation | Add list data validation to a merged range in Excel with Aspose.Cells .NET | Aspose.Cells example for merging cells and setting a validation list in C# | How to apply a dropdown list to merged cells using Aspose.Cells in a .NET application | Saving an Excel workbook after merging cells and adding data validation with Aspose.Cells C#
// Tags: merge cells Aspose.Cells C# | list validation merged range Aspose.Cells | Excel dropdown validation Aspose.Cells .NET | save workbook after applying validation Aspose.Cells | cellarea merge validation example Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing workbook (or creates a new one), merges cells D2:E2 on the first worksheet, attaches a list‑type data validation with three options to the merged range, and saves the updated file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // default workbook with one worksheet
            }

            Worksheet sheet = workbook.Worksheets[0];

            // Define the merged range D2:E2 (zero‑based indexes).
            CellArea mergeArea = new CellArea
            {
                StartRow = 1,    // Row 2
                StartColumn = 3, // Column D
                EndRow = 1,      // Row 2
                EndColumn = 4    // Column E
            };

            // Merge the cells.
            sheet.Cells.Merge(
                mergeArea.StartRow,
                mergeArea.StartColumn,
                mergeArea.EndRow - mergeArea.StartRow + 1,
                mergeArea.EndColumn - mergeArea.StartColumn + 1);

            // Add a list‑type data validation for the merged range.
            int validationIndex = sheet.Validations.Add(mergeArea);
            Validation validation = sheet.Validations[validationIndex];
            validation.Type = ValidationType.List;
            validation.Formula1 = "\"Option1,Option2,Option3\"";
            validation.ShowError = true;
            validation.ErrorTitle = "Invalid entry";
            validation.ErrorMessage = "Please select a value from the list.";

            // Save the workbook with the changes.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
