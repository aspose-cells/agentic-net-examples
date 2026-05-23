using System;
using Aspose.Cells;

namespace AsposeCellsVariableDemo
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook (template)
            Workbook workbook = new Workbook();

            // 2. Add a worksheet that will hold variables (smart marker variables)
            Worksheet variablesSheet = workbook.Worksheets.Add("Variables");
            // Define a variable name in column A and its value in column B
            variablesSheet.Cells["A1"].PutValue("ReportDate");   // variable name
            variablesSheet.Cells["B1"].PutValue(DateTime.Today); // variable value

            // 3. Create a WorkbookDesigner and assign the workbook
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // 4. Tell the designer which worksheet contains the variables
            designer.VariablesWorksheetName = "Variables";

            // 5. Add a template worksheet that uses the variable via a smart marker
            Worksheet templateSheet = workbook.Worksheets.Add("Template");
            // Smart marker that references the variable defined above
            templateSheet.Cells["A1"].PutValue("&=$ReportDate");

            // 6. Process the smart markers – the variable will be replaced with its value
            designer.Process();

            // 7. Save the resulting workbook
            workbook.Save("VariableDemo.xlsx");
        }
    }
}