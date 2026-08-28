// Title: Generate an Excel file with a dynamic dropdown list populated from a DataTable using Aspose.Cells smart markers (C#)
// AI Prompts: Create a workbook where a data sheet is filled using smart markers from a DataTable and apply a list‑type validation on another sheet that references the filled column. | Change the example to use the Value column as the dropdown source and move the validation to cell B5, keeping the same WorkbookDesigner workflow.
// Common Searches: how to use Aspose.Cells WorkbookDesigner to fill a sheet and create a dropdown list in C# | dynamic Excel data validation list from DataTable using smart markers Aspose.Cells | populate Excel dropdown with values from a DataTable via smart markers in C# example
// Tags: Aspose.Cells smart markers populate reference sheet | list validation range from smart marker output | dynamic dropdown from DataTable Aspose.Cells | WorkbookDesigner set data source for Excel validation | C# generate Excel dropdown using smart markers

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDynamicDropdown
{
    // The sample builds a template workbook, uses a DataTable as a smart‑marker data source to fill a reference worksheet, creates a list‑type data validation on Form!A2 that points to the populated Category column, and saves the file as DynamicDropdown.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // -------------------- 1. Create template workbook --------------------
                Workbook template = new Workbook();

                // Sheet that will hold the reference data (populated via smart markers)
                Worksheet dataSheet = template.Worksheets[0];
                dataSheet.Name = "ReferenceData";

                // Header row
                dataSheet.Cells["A1"].PutValue("Category");
                dataSheet.Cells["B1"].PutValue("Value");

                // Smart markers for data rows (will be repeated for each DataTable row)
                dataSheet.Cells["A2"].PutValue("&=RefTable.Category");
                dataSheet.Cells["B2"].PutValue("&=RefTable.Value");

                // Sheet where the dropdown will appear
                Worksheet formSheet = template.Worksheets.Add("Form");
                // Target cell for dropdown
                formSheet.Cells["A2"].PutValue("Select Category");

                // -------------------- 2. Prepare data source --------------------
                DataTable dt = new DataTable("RefTable");
                dt.Columns.Add("Category", typeof(string));
                dt.Columns.Add("Value", typeof(int));

                dt.Rows.Add("Apple", 10);
                dt.Rows.Add("Banana", 20);
                dt.Rows.Add("Cherry", 30);
                dt.Rows.Add("Date", 40);

                // -------------------- 3. Process smart markers --------------------
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = template
                };
                designer.SetDataSource(dt); // DataTable name matches smart marker table name
                designer.Process(); // Populate the reference data

                // -------------------- 4. Create dynamic dropdown list --------------------
                // Determine the range that now contains the populated categories
                int lastDataRow = dataSheet.Cells.MaxDataRow; // zero‑based index, includes header
                // Build address like: ReferenceData!$A$2:$A${lastDataRow+1}
                string listRange = $"ReferenceData!$A$2:$A${lastDataRow + 1}";

                // Add validation to the target cell (Form!A2) using CellArea overload
                CellArea targetArea = CellArea.CreateCellArea("A2", "A2");
                int validationIndex = formSheet.Validations.Add(targetArea);
                Validation validation = formSheet.Validations[validationIndex];
                validation.Type = ValidationType.List;   // set validation type
                validation.Formula1 = listRange;          // reference to the populated list
                validation.InCellDropDown = true;         // enable dropdown arrow

                // -------------------- 5. Save result --------------------
                string outputPath = "DynamicDropdown.xlsx";
                designer.Workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
