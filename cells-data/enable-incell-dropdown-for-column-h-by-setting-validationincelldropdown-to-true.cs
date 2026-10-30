// Title: Enable an in‑cell dropdown list for column H in an Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells add list validation dropdown to column H | how to set Validation.InCellDropDown true for a range using Aspose.Cells | create Excel dropdown list in column H with Aspose.Cells .NET | Aspose.Cells example for in‑cell dropdown list validation | define CellArea H1:H100 and enable dropdown in Aspose.Cells
// Tags: set Validation.InCellDropDown property | list validation for column H | Aspose.Cells create Excel dropdown list | C# define CellArea range | list validation formula with comma‑separated items

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample creates a new workbook, defines the range H1:H100 using a CellArea, adds a list validation to that range, sets the validation type to List, enables the in‑cell dropdown by setting Validation.InCellDropDown to true, supplies the dropdown items via the Formula1 property, and saves the file as InCellDropdown.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Define the range for column H (H1:H100)
                CellArea area = new CellArea
                {
                    StartRow = 0,
                    EndRow = 99,
                    StartColumn = 7,
                    EndColumn = 7
                };

                // Add a list validation to the defined range
                int validationIndex = sheet.Validations.Add(area);
                Validation validation = sheet.Validations[validationIndex];

                // Set validation type to List and enable in‑cell dropdown
                validation.Type = ValidationType.List;
                validation.InCellDropDown = true;

                // Define list items for the dropdown
                validation.Formula1 = "\"Option1,Option2,Option3\"";

                // Save the workbook
                string outputPath = "InCellDropdown.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
