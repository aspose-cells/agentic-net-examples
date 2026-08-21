// Title: Create a drop‑down list using a named range with Aspose.Cells for .NET (C#)
// Description: C# example that builds a workbook, fills A1:A5 with items, creates a named range "MyList", adds a list‑type data validation to cell B1 referencing the named range, enables an in‑cell drop‑down, and saves the file.
// Keywords: Aspose.Cells | C# | named range | data validation | drop-down list | list validation | Excel automation | Workbook | Worksheet
// Common Searches: Aspose.Cells named range validation C# | How to create drop down list from named range in Aspose.Cells | C# data validation list using named range Aspose.Cells | Set in‑cell dropdown with named range Aspose.Cells .NET | Add list validation to a cell Aspose.Cells
// Developer Intent: Define a named range and use it as the source for a list‑type data validation (drop‑down) in another cell.
// Use Cases: Provide users with a selectable list of predefined options next to a data entry column. | Build reusable Excel templates where validation lists are driven by named ranges for easy maintenance. | Generate multiple workbooks programmatically, each with dynamic drop‑down lists based on named ranges.
// AI Prompts: Show how to reference a named range located on a different worksheet when creating list validation. | Provide code that creates several named ranges and assigns each to a different cell’s drop‑down list. | Explain how to update the contents of a named range after the workbook is saved and ensure the dropdown reflects the changes.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // C# example that builds a workbook, fills A1:A5 with items, creates a named range "MyList", adds a list‑type data validation to cell B1 referencing the named range, enables an in‑cell drop‑down, and saves the file.
    public class NamedRangeValidationDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1";

            // Populate sample values for the named range (A1:A5)
            for (int i = 0; i < 5; i++)
            {
                sheet.Cells[i, 0].PutValue("Item " + (i + 1));
            }

            // Add a named range that refers to the list items (A1:A5)
            int nameIndex = workbook.Worksheets.Names.Add("MyList");
            workbook.Worksheets.Names[nameIndex].RefersTo = "=Sheet1!$A$1:$A$5";

            // Create a validation for cell B1 using the named range as its list source
            Validation validation = sheet.Cells["B1"].GetValidation();
            validation.Type = ValidationType.List;
            validation.Formula1 = "MyList"; // Named range name
            validation.InCellDropDown = true;
            validation.InputMessage = "Select an item from the list.";
            validation.InputTitle = "Choose Item";

            // Save the workbook
            string outputPath = "NamedRangeValidationDemo.xlsx";
            workbook.Save(outputPath);
        }
    }
}
