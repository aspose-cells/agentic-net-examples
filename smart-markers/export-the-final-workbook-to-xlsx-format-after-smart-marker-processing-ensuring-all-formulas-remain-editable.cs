using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace SmartMarkerExportDemo
{
    class Program
    {
        static void Main()
        {
            // Load the template workbook that contains smart markers
            Workbook workbook = new Workbook("template.xlsx");

            // Initialize the WorkbookDesigner with the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };

            // Example: set a JSON data source for smart markers
            string jsonData = @"{""Name"":""Sample Product"",""Price"":123.45}";
            designer.SetJsonDataSource("Data", jsonData);

            // Process all smart markers in the workbook
            designer.Process();

            // Ensure formulas are not forced to calculate on save
            // (formulas will stay as editable expressions in the output file)
            workbook.Settings.FormulaSettings.CalculateOnSave = false;

            // Save the processed workbook to XLSX format
            // The Save(string) overload infers the format from the file extension
            workbook.Save("output.xlsx");
        }
    }
}