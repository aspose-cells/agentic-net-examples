// Title: Create an in‑cell dropdown list for rows 1‑100 in column H using Aspose.Cells for .NET
// AI Prompts: Generate C# code that defines a CellArea covering rows 1‑100 in column H, adds a List‑type validation with the values Option1, Option2, Option3, sets Validation.InCellDropDown to true, and saves the workbook. | Write a C# snippet using Aspose.Cells to apply an in‑cell dropdown to column H (rows 1‑100) by configuring Validation.Type = List and enabling the dropdown.
// Common Searches: Aspose.Cells C# add dropdown list to column H rows 1 to 100 | How to enable in‑cell dropdown for a specific column using Aspose.Cells .NET | Set list validation with InCellDropDown property in Aspose.Cells workbook | Create Excel dropdown in column H with Aspose.Cells C# example | Aspose.Cells validation for column H range 1‑100 list values
// Tags: Aspose.Cells list validation column H | C# Aspose.Cells enable in‑cell dropdown | Aspose.Cells Validation.InCellDropDown property | Define CellArea for column range Aspose.Cells | Excel dropdown list generation Aspose.Cells C#

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example creates a new workbook, defines a CellArea covering rows 1‑100 in column H, adds a list‑type validation with three options, enables the InCellDropDown property, and saves the file as ColumnH_InCellDropdown.xlsx.
    public class EnableInCellDropdownForColumnH
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Define the validation area for column H (index 7), rows 1 to 100 (0‑based indices)
                CellArea area = new CellArea
                {
                    StartRow = 0,      // Row 1
                    EndRow = 99,       // Row 100
                    StartColumn = 7,   // Column H
                    EndColumn = 7
                };

                // Add a new validation to the worksheet for the defined area
                int validationIndex = worksheet.Validations.Add(area);
                Validation validation = worksheet.Validations[validationIndex];

                // Set the validation type to List and provide the list of acceptable values
                validation.Type = ValidationType.List;
                validation.Formula1 = "Option1,Option2,Option3";

                // Enable the in‑cell drop‑down list
                validation.InCellDropDown = true;

                // Save the workbook to a file
                workbook.Save("ColumnH_InCellDropdown.xlsx");
                Console.WriteLine("Workbook saved successfully as ColumnH_InCellDropdown.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            EnableInCellDropdownForColumnH.Run();
        }
    }
}
