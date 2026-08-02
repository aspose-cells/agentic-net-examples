using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerValidation
{
    class Program
    {
        static void Main()
        {
            // Load the template workbook that contains smart markers
            Workbook workbook = new Workbook("TemplateWithSmartMarkers.xlsx");

            // Access the first worksheet (where smart markers are placed)
            Worksheet sheet = workbook.Worksheets[0];

            // Prepare a data source for the smart markers
            List<Product> products = new List<Product>
            {
                new Product { Name = "Apple", Category = "Fruit" },
                new Product { Name = "Carrot", Category = "Vegetable" },
                new Product { Name = "Milk", Category = "Dairy" }
            };

            // Set up the WorkbookDesigner, assign the workbook and data source
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook,
                LineByLine = false
            };
            designer.SetDataSource("Products", products);

            // Process all smart markers in the workbook
            designer.Process();

            // After processing, apply data validation to the "Category" column
            // Assuming the Category values are placed in column B (index 1) starting from row 2
            int startRow = 1; // zero‑based index for row 2
            int endRow = startRow + products.Count - 1;
            CellArea validationArea = new CellArea
            {
                StartRow = startRow,
                StartColumn = 1,
                EndRow = endRow,
                EndColumn = 1
            };

            // Create a list‑type validation with a drop‑down
            ValidationCollection validations = sheet.Validations;
            int validationIndex = validations.Add(validationArea);
            Validation validation = validations[validationIndex];
            validation.Type = ValidationType.List;
            validation.InCellDropDown = true;
            validation.ShowInput = true;
            validation.InputTitle = "Category";
            validation.InputMessage = "Select a valid category.";
            validation.ErrorTitle = "Invalid Category";
            validation.ErrorMessage = "Please select a category from the list.";
            validation.ShowError = true;
            // Define the allowed list values directly
            validation.Formula1 = "Fruit,Vegetable,Dairy";

            // Save the workbook with the applied validation
            workbook.Save("OutputWithValidation.xlsx");
        }

        // Simple POCO class representing product data
        public class Product
        {
            public string Name { get; set; }
            public string Category { get; set; }
        }
    }
}