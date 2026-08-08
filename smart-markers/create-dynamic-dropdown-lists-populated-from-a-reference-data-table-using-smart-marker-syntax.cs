// Title: C# – Create a Dynamic Dropdown List from a Smart‑Marker Data Source with Aspose.Cells
// Description: This example shows how to build an in‑memory Excel template, populate a column with a smart‑marker (&=Colors.ColorName) using WorkbookDesigner, calculate the resulting range, and attach a list‑type data‑validation dropdown to cell B2 that automatically reflects the populated values. The workbook is saved as DynamicDropdown.xlsx.
// Keywords: Aspose.Cells | C# | dynamic dropdown | smart markers | WorkbookDesigner | data validation list | Excel dropdown from DataTable | populate Excel list | in‑memory workbook template | .NET Excel automation
// Common Searches: Aspose.Cells create dropdown from smart marker C# | add data validation list after WorkbookDesigner processing | dynamic Excel dropdown using DataTable and smart markers | C# generate Excel template with smart markers and dropdown | populate Excel list validation range programmatically
// Developer Intent: Generate an Excel file where a dropdown in cell B2 is automatically filled with values supplied by a DataTable through a smart‑marker expansion.
// Use Cases: Design a reporting template that lists database values and lets users select an item from a dropdown that grows with the source data. | Build a reusable Excel form where the validation list updates automatically whenever new rows are added to the underlying DataTable. | Create a configurable spreadsheet where end‑users can pick options that are derived from dynamic business data.
// AI Prompts: Write C# code using Aspose.Cells to add a list‑type data validation to a cell, referencing a range populated by a smart marker. | Explain how to find the last populated row after WorkbookDesigner processes smart markers and use it to set the validation formula. | Show how to combine a DataTable as a data source with a smart marker and then attach a dropdown validation to another cell in the same worksheet.

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// This example shows how to build an in‑memory Excel template, populate a column with a smart‑marker (&=Colors.ColorName) using WorkbookDesigner, calculate the resulting range, and attach a list‑type data‑validation dropdown to cell B2 that automatically reflects the populated values. The workbook is saved as DynamicDropdown.xlsx.
class DynamicDropdownDemo
{
    static void Main()
    {
        try
        {
            // ------------------------------------------------------------
            // 1. Create a template workbook in memory
            // ------------------------------------------------------------
            Workbook template = new Workbook();
            Worksheet ws = template.Worksheets[0];
            Cells cells = ws.Cells;

            // Header for the reference list
            cells["A1"].PutValue("Colors");

            // Smart marker that will be replaced by the data source rows
            // The smart marker populates column A starting from A2
            cells["A2"].PutValue("&=Colors.ColorName");

            // Cell where the dropdown will be placed (B2)
            cells["B2"].PutValue("Select Color");

            // ------------------------------------------------------------
            // 2. Save the template to a memory stream (required by the rule)
            // ------------------------------------------------------------
            using (MemoryStream templateStream = new MemoryStream())
            {
                template.Save(templateStream, SaveFormat.Xlsx);
                templateStream.Position = 0; // reset for reading

                // ------------------------------------------------------------
                // 3. Prepare the reference data table
                // ------------------------------------------------------------
                DataTable colorsTable = new DataTable("Colors");
                colorsTable.Columns.Add("ColorName", typeof(string));
                colorsTable.Rows.Add("Red");
                colorsTable.Rows.Add("Green");
                colorsTable.Rows.Add("Blue");
                colorsTable.Rows.Add("Yellow");

                // ------------------------------------------------------------
                // 4. Initialize WorkbookDesigner with the template and set data source
                // ------------------------------------------------------------
                WorkbookDesigner designer = new WorkbookDesigner();
                designer.Workbook = new Workbook(templateStream);
                designer.SetDataSource(colorsTable);
                designer.Process(); // populate the smart marker

                // ------------------------------------------------------------
                // 5. After processing, add data‑validation dropdown that points to the populated list
                // ------------------------------------------------------------
                Worksheet sheet = designer.Workbook.Worksheets[0];

                // Determine the last row that contains data in column A (the list)
                int lastDataRow = sheet.Cells.MaxDataRow; // includes header row

                // Build the address of the list range (exclude header)
                string listRange = $"A2:A{lastDataRow}";

                // Define the cell area for the dropdown (B2)
                CellArea dropdownArea = new CellArea
                {
                    StartRow = 1,    // B2 row (0‑based)
                    StartColumn = 1, // B column
                    EndRow = 1,
                    EndColumn = 1
                };

                // Add validation for the target cell using the overload that accepts CellArea
                int validationIndex = sheet.Validations.Add(dropdownArea);
                Validation validation = sheet.Validations[validationIndex];
                validation.Type = ValidationType.List;
                validation.Formula1 = listRange;          // reference the list range
                validation.InCellDropDown = true;        // enable the dropdown arrow

                // ------------------------------------------------------------
                // 6. Save the final workbook
                // ------------------------------------------------------------
                string outputPath = "DynamicDropdown.xlsx";

                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                designer.Workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
