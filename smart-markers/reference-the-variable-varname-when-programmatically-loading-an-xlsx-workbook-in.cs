using System;
using Aspose.Cells;

namespace AsposeCellsVariableReferenceDemo
{
    public class Program
    {
        public static void Main()
        {
            // Name of the variable to be referenced in the workbook
            string varName = "MyVariable";

            // Create a new workbook and add a smart marker
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";
            // Place a smart marker in cell A1
            sheet.Cells["A1"].PutValue($"&={{${varName}}}");

            // Initialize WorkbookDesigner with the created workbook
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Set the value for the variable referenced in the workbook
            designer.SetDataSource(varName, "Hello, Aspose!");

            // Process the smart markers to replace them with the variable value
            designer.Process();

            // Save the processed workbook
            string outputPath = "ProcessedWorkbook.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook processed and saved to '{outputPath}'. Variable '{varName}' was applied.");
        }
    }
}