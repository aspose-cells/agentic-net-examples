// Title: Merge cells D2:E2, add a list‑type data validation dropdown, and save the workbook with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# merge cells D2:E2 and create a dropdown list validation | C# code to add list data validation to a merged cell using Aspose.Cells | Save changes after merging cells and applying data validation with Aspose.Cells .NET | How to set input message for Excel data validation using Aspose.Cells C# | Create a dropdown list in cell D2 after merging D2:E2 with Aspose.Cells
// Tags: merge cells with Aspose.Cells C# | list data validation dropdown Aspose.Cells | cell area validation API Aspose.Cells | save workbook to file Aspose.Cells .NET | input message for Excel validation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file or creates a new workbook, merges cells D2 and E2, defines a validation area for cell D2, adds a list‑type data validation with three options and an input prompt, and saves the modified workbook to output.xlsx, handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
            }

            Worksheet sheet = workbook.Worksheets[0];

            // Merge cells D2:E2 (zero‑based indexes: column D = 3, row 2 = 1)
            sheet.Cells.Merge(3, 1, 1, 2);

            // Define the area for validation (cell D2)
            CellArea area = new CellArea
            {
                StartRow = 1,
                StartColumn = 3,
                EndRow = 1,
                EndColumn = 3
            };

            // Add validation to the defined area
            int validationIndex = sheet.Validations.Add(area);
            Validation validation = sheet.Validations[validationIndex];

            // Set validation type to List with comma‑separated items
            validation.Type = ValidationType.List;
            validation.Formula1 = "\"Option1,Option2,Option3\"";
            validation.InCellDropDown = true;

            // Show input message (prompt) when the cell is selected
            validation.InputTitle = "Select Option";
            validation.InputMessage = "Please choose an item from the list.";

            // Save the modified workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
