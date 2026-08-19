// Title: Add Data Validation to Smart‑Marker‑Generated Cells with Aspose.Cells for .NET (C#)
// Description: Loads a template workbook, binds a List<Person> as the smart‑marker data source, processes the markers, then defines a CellArea for the age column and applies a whole‑number validation (18‑65) before saving the file.
// Keywords: Aspose.Cells | smart markers | data validation | C# | .NET | whole number validation | age range | CellArea | ValidationCollection | WorkbookDesigner | Excel automation
// Common Searches: Aspose.Cells add validation after smart markers | C# data validation for smart‑marker output | restrict age column in generated Excel using Aspose | how to apply whole number validation in Aspose.Cells | smart markers validation example .NET
// Developer Intent: Apply a numeric range rule to cells that were filled by smart markers after the workbook has been processed.
// Use Cases: Guarantee ages generated from a Person list stay between 18 and 65, preventing manual entry errors. | Enforce salary or score limits in any column populated by smart markers. | Lock down user edits in automatically created reports to maintain data integrity. | Add validation to multiple smart‑marker columns with a single post‑processing step.
// AI Prompts: Generate C# code to add a date range validation to a column populated by smart markers in Aspose.Cells. | Show how to create a dropdown list validation from an enum for smart‑marker‑filled cells. | Explain how to calculate validation limits dynamically from the size of the data source when using Aspose.Cells smart markers. | Provide an example of applying a custom formula validation to smart‑marker output. | Write code that sets validation for several columns after processing smart markers.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerValidation
{
    // Loads a template workbook, binds a List<Person> as the smart‑marker data source, processes the markers, then defines a CellArea for the age column and applies a whole‑number validation (18‑65) before saving the file.
    public class Program
    {
        public static void Main()
        {
            // Load the template workbook that contains smart markers
            Workbook workbook = new Workbook("TemplateWithSmartMarkers.xlsx");

            // Initialize WorkbookDesigner with the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Prepare sample data source for smart markers
            List<Person> persons = new List<Person>
            {
                new Person { Name = "Alice", Age = 28 },
                new Person { Name = "Bob", Age = 35 },
                new Person { Name = "Charlie", Age = 42 }
            };

            // Set the data source (the name must match the smart marker prefix in the template)
            designer.SetDataSource("Persons", persons);

            // Process all smart markers in the workbook
            designer.Process();

            // After processing, apply data validation to the cells that were filled by the smart markers.
            // Assume the template placed ages in column B starting from row 2 (B2:B4).
            Worksheet sheet = workbook.Worksheets[0];

            // Define the validation area (B2:B4)
            CellArea ageArea = CellArea.CreateCellArea(1, 1, persons.Count, 1); // rows are zero‑based

            // Add a new validation to the worksheet
            ValidationCollection validations = sheet.Validations;
            int validationIndex = validations.Add(ageArea);
            Validation ageValidation = validations[validationIndex];

            // Configure the validation: whole numbers between 18 and 65
            ageValidation.Type = ValidationType.WholeNumber;
            ageValidation.Operator = OperatorType.Between;
            ageValidation.Formula1 = "18";
            ageValidation.Formula2 = "65";
            ageValidation.InputTitle = "Age Input";
            ageValidation.InputMessage = "Enter an age between 18 and 65.";
            ageValidation.ErrorTitle = "Invalid Age";
            ageValidation.ErrorMessage = "The age must be a whole number between 18 and 65.";
            ageValidation.ShowInput = true;
            ageValidation.ShowError = true;
            ageValidation.InCellDropDown = false;
            ageValidation.IgnoreBlank = true;

            // Save the resulting workbook
            workbook.Save("OutputWithValidation.xlsx");
        }
    }

    // Simple POCO class used as a data source for smart markers
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }
}
