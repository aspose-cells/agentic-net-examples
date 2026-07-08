using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // 1. Create a new workbook
        Workbook workbook = new Workbook();

        // 2. Add a worksheet that will hold variables (smart marker variables)
        Worksheet variablesSheet = workbook.Worksheets.Add("Variables");
        // Define a variable named "Variable1" with the value "Hello Aspose"
        variablesSheet.Cells["A1"].PutValue("Variable1");
        variablesSheet.Cells["B1"].PutValue("Hello Aspose");

        // 3. Create a WorkbookDesigner instance and assign the workbook
        WorkbookDesigner designer = new WorkbookDesigner(workbook);

        // 4. Tell the designer which worksheet contains the variables
        designer.VariablesWorksheetName = "Variables";

        // 5. Add a template worksheet that uses the variable via a smart marker
        Worksheet templateSheet = workbook.Worksheets.Add("Template");
        // Smart marker that references the variable defined above
        templateSheet.Cells["A1"].PutValue("&=$Variable1");

        // 6. Process the smart markers – the variable value will be inserted
        designer.Process();

        // 7. Save the resulting workbook
        workbook.Save("VariableDemo.xlsx");
    }
}