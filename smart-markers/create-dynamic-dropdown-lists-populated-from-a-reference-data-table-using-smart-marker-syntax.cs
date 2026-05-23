using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsDynamicDropdown
{
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // 2. Prepare reference data table (the source for the dropdown list)
                DataTable dtColors = new DataTable("Colors");
                dtColors.Columns.Add("ColorName", typeof(string));
                dtColors.Rows.Add("Red");
                dtColors.Rows.Add("Green");
                dtColors.Rows.Add("Blue");
                dtColors.Rows.Add("Yellow");

                // 3. Define a smart marker that will be replaced by the data source values
                cells["A2"].PutValue("&=Colors.ColorName");

                // 4. Create a named range that will hold the list items after processing
                AsposeRange listRange = sheet.Cells.CreateRange("A2:A5");
                listRange.Name = "ColorsList";

                // 5. Add a cell (B2) where the dropdown will appear and set data validation
                Validation validation = cells["B2"].GetValidation();
                validation.Type = ValidationType.List;
                validation.Formula1 = "=ColorsList";   // reference the named range
                validation.InCellDropDown = true;     // enable the dropdown arrow

                // 6. Bind the data source and process the smart markers
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };
                designer.SetDataSource(dtColors); // Table name "Colors" is taken from DataTable.TableName
                designer.Process();               // Populate the smart marker range with actual values

                // 7. Save the resulting workbook
                string outputPath = "DynamicDropdown.xlsx";
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