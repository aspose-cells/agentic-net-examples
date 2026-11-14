// Title: Add integer range and drop-down list data validation to smart-marker generated columns using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add whole number validation to a column created by WorkbookDesigner | How to set a drop-down list validation on smart-marker generated cells in Excel with Aspose | Restrict user input to 1-100 in a smart-marker column using Aspose.Cells validation API | Apply data validation after processing smart markers in a .NET workbook | C# Aspose.Cells validation for smart marker range B2:B100 list values
// Tags: Aspose.Cells data validation whole number | Aspose.Cells list validation smart markers | WorkbookDesigner post-processing validation C# | Excel column integer range validation Aspose | smart marker generated column validation .NET

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.Data;
using System.IO;

// The example loads or creates an Excel template, defines a DataSet with an Items table, and processes smart markers via WorkbookDesigner. After processing, it adds a whole-number validation (1‑100) to the Quantity column (C2:C100) and a drop‑down list validation (Item1,Item2,Item3) to the Name column (B2:B100), then saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string templatePath = "Template.xlsx";
            const string outputPath = "Output.xlsx";

            // Load the template workbook if it exists; otherwise create a new workbook.
            Workbook workbook;
            if (File.Exists(templatePath))
            {
                workbook = new Workbook(templatePath);
            }
            else
            {
                workbook = new Workbook();
                // Optionally add headers for smart markers (e.g., &=Items.Name, &=Items.Quantity)
                Worksheet ws = workbook.Worksheets[0];
                ws.Cells["B2"].PutValue("&=Items.Name");
                ws.Cells["C2"].PutValue("&=Items.Quantity");
            }

            // Prepare data source as a DataSet (required by WorkbookDesigner).
            DataTable table = new DataTable("Items");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Quantity", typeof(int));

            table.Rows.Add("Item1", 10);
            table.Rows.Add("Item2", 20);

            DataSet dataSet = new DataSet();
            dataSet.Tables.Add(table);

            // Process smart markers.
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dataSet);
            designer.Process();

            // Apply data validation after processing.
            Worksheet sheet = workbook.Worksheets[0];

            // -----------------------------------------------------------------
            // Example 1: Restrict the "Quantity" column (C2:C100) to integers 1‑100
            // -----------------------------------------------------------------
            int firstRow = 1;   // Row 2 (zero‑based)
            int lastRow = 99;   // Row 100
            int quantityCol = 2; // Column C (zero‑based)

            CellArea quantityArea = new CellArea
            {
                StartRow = firstRow,
                StartColumn = quantityCol,
                EndRow = lastRow,
                EndColumn = quantityCol
            };

            int qtyValidationIdx = sheet.Validations.Add(quantityArea);
            Validation qtyValidation = sheet.Validations[qtyValidationIdx];
            qtyValidation.Type = ValidationType.WholeNumber; // integer validation
            qtyValidation.Operator = OperatorType.Between;
            qtyValidation.Formula1 = "1";
            qtyValidation.Formula2 = "100";
            qtyValidation.ShowError = true;
            qtyValidation.ErrorTitle = "Invalid Quantity";
            qtyValidation.ErrorMessage = "Please enter a number between 1 and 100.";

            // -----------------------------------------------------------------
            // Example 2: Restrict the "Name" column (B2:B100) to a drop‑down list
            // -----------------------------------------------------------------
            int nameCol = 1; // Column B (zero‑based)

            CellArea nameArea = new CellArea
            {
                StartRow = firstRow,
                StartColumn = nameCol,
                EndRow = lastRow,
                EndColumn = nameCol
            };

            int nameValidationIdx = sheet.Validations.Add(nameArea);
            Validation nameValidation = sheet.Validations[nameValidationIdx];
            nameValidation.Type = ValidationType.List;
            // List values must be quoted and comma‑separated
            nameValidation.Formula1 = "\"Item1,Item2,Item3\"";
            nameValidation.ShowError = true;
            nameValidation.ErrorTitle = "Invalid Name";
            nameValidation.ErrorMessage = "Select a name from the list.";

            // Save the workbook with the applied validations.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
