// Title: Generate an Excel file with a smart‑marker driven dropdown list using Aspose.Cells in C#
// AI Prompts: Write C# code that creates a reference worksheet, loads categories into a DataTable, replaces a ${Category} smart marker with WorkbookDesigner, and attaches a list‑type validation to cell B1. | Show how to set up ValidationCollection so the dropdown list reads its items from a range on the Data sheet. | Provide a complete example that saves the workbook as DynamicDropdown.xlsx and includes basic exception handling.
// Common Searches: C# Aspose.Cells create dropdown list that references another sheet | how to use smart markers with data validation in Aspose.Cells | Aspose.Cells example for populating a list validation from a DataTable | generate Excel file with dynamic dropdown using Aspose.Cells designer | Excel dropdown list from Data sheet using Aspose.Cells C#
// Tags: Aspose.Cells WorkbookDesigner smart marker processing | list validation referencing external worksheet Aspose.Cells | C# generate Excel dropdown via ValidationCollection | populate dropdown values from DataTable Aspose.Cells | dynamic list validation implementation Aspose.Cells

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The sample creates a new workbook, adds a 'Data' sheet with a category list, binds the list to a DataTable, processes a ${Category} smart marker on a 'Form' sheet using WorkbookDesigner, then applies a list‑type data validation to cell B1 that points to the category range on the Data sheet, and finally saves the file as DynamicDropdown.xlsx.
class DynamicDropdownWithSmartMarkers
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // 1. Create reference data table on sheet "Data"
            // -------------------------------------------------
            Worksheet dataSheet = workbook.Worksheets[workbook.Worksheets.Add()];
            dataSheet.Name = "Data";

            // Header
            dataSheet.Cells["A1"].PutValue("Category");

            // Sample reference data
            string[] categories = { "Fruit", "Vegetable", "Dairy", "Meat", "Grain" };
            for (int i = 0; i < categories.Length; i++)
            {
                dataSheet.Cells[i + 1, 0].PutValue(categories[i]); // A2, A3, ...
            }

            // -------------------------------------------------
            // 2. Prepare the form sheet with a smart marker
            // -------------------------------------------------
            Worksheet formSheet = workbook.Worksheets[workbook.Worksheets.Add()];
            formSheet.Name = "Form";

            // Header for the dropdown column
            formSheet.Cells["A1"].PutValue("Select Category");

            // Smart marker that will be replaced by a value from the data source
            formSheet.Cells["B1"].PutValue("${Category}");

            // -------------------------------------------------
            // 3. Set up the data source for the smart marker
            // -------------------------------------------------
            DataTable dt = new DataTable();
            dt.Columns.Add("Category", typeof(string));
            foreach (string cat in categories)
            {
                dt.Rows.Add(cat);
            }

            // Process the smart marker
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dt);
            designer.Process(); // Replaces ${Category} with the first row value

            // -------------------------------------------------
            // 4. Add a dynamic dropdown list to the cell B1
            // -------------------------------------------------
            // Define the range for the dropdown list (e.g., Data!A2:A6)
            int firstRow = 2; // Excel rows are 1‑based, A2 is row 2
            int lastRow = firstRow + categories.Length - 1;
            string listRange = $"Data!$A${firstRow}:$A${lastRow}";

            ValidationCollection validations = formSheet.Validations;

            // Create a CellArea for B1 (row 0, column 1)
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 1,
                EndRow = 0,
                EndColumn = 1
            };

            // Add validation for the specified area
            validations.Add(area);
            Validation validation = validations[validations.Count - 1];
            validation.Type = ValidationType.List;
            validation.Operator = OperatorType.None;
            validation.Formula1 = $"={listRange}";
            validation.InCellDropDown = true;
            validation.ShowError = true;
            validation.ErrorTitle = "Invalid selection";
            validation.ErrorMessage = "Please select a value from the list.";

            // -------------------------------------------------
            // 5. Save the workbook
            // -------------------------------------------------
            string outputPath = "DynamicDropdown.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
