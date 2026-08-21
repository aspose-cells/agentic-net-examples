// Title: C# – Add Column to Aspose.Cells Table and Apply List Validation from Another Column
// Description: Shows how to create a workbook, define a ListObject for A1:B4, insert a new header, expand the table to include the extra column, and configure an in‑cell drop‑down list for the new cells using the values from the Category column (B2:B4). The file is saved as TableWithValidation.xlsx.
// Keywords: Aspose.Cells C# add column to table | Aspose.Cells list validation from column | ListObject resize Aspose.Cells | in‑cell drop‑down Aspose.Cells | C# workbook table validation | TableWithValidation.xlsx example | Aspose.Cells data validation list
// Common Searches: add column to Aspose.Cells ListObject C# | Aspose.Cells set data validation list from another column | resize Aspose.Cells table after adding column | C# create drop‑down list validation in table | Aspose.Cells example for table validation
// Developer Intent: Expand an existing ListObject with an extra column and attach a list‑type validation that references a different column on the same sheet.
// Use Cases: Build a product sheet where the new "Option" column offers selectable values drawn from the "Category" column. | Generate a template that programmatically grows a table and enforces consistent entries via a drop‑down list. | Create a reporting form that lets users pick values from a dynamically maintained source list.
// AI Prompts: Write C# code with Aspose.Cells to insert a column into a ListObject, resize the table, and add a list validation that pulls items from another column. | Explain step‑by‑step how to set up an in‑cell drop‑down validation after expanding an Aspose.Cells table. | Provide guidance on linking a validation list to a column range within the same worksheet using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;   // Required for ListObject

namespace AsposeCellsTableValidationDemo
{
    // Shows how to create a workbook, define a ListObject for A1:B4, insert a new header, expand the table to include the extra column, and configure an in‑cell drop‑down list for the new cells using the values from the Category column (B2:B4). The file is saved as TableWithValidation.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the initial table (Item and Category columns)
                sheet.Cells["A1"].PutValue("Item");
                sheet.Cells["B1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["A4"].PutValue("Cherry");
                sheet.Cells["B2"].PutValue("Fruit");
                sheet.Cells["B3"].PutValue("Fruit");
                sheet.Cells["B4"].PutValue("Fruit");

                // Add a ListObject (table) that covers the range A1:B4
                int tableIdx = sheet.ListObjects.Add("A1", "B4", true);
                ListObject table = sheet.ListObjects[tableIdx];

                // Add a new column header for the column that will hold the validation list
                sheet.Cells["C1"].PutValue("Option");

                // Expand the table to include the new column (C)
                // New area: rows 0‑3 (A‑D), columns 0‑2 (A‑C)
                // Resize(firstRow, firstColumn, totalRows, totalColumns, expandRows)
                table.Resize(0, 0, 4, 3, false);

                // Define the validation area: column C (index 2), rows 2‑4 (indexes 1‑3)
                CellArea validationArea = CellArea.CreateCellArea(1, 2, 3, 2);

                // Add a validation to the defined area
                int validationIdx = sheet.Validations.Add(validationArea);
                Validation validation = sheet.Validations[validationIdx];

                // Configure the validation as a drop‑down list sourced from the Category column (B2:B4)
                validation.Type = ValidationType.List;
                validation.InCellDropDown = true;
                validation.Formula1 = "$B$2:$B$4";

                // Save the workbook
                workbook.Save("TableWithValidation.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
