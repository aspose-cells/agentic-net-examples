using System;
using Aspose.Cells;

class DefineVariableDemo
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        string variablesSheetName = "Variables";
        Worksheet variablesSheet = workbook.Worksheets[variablesSheetName];
        if (variablesSheet == null)
        {
            workbook.Worksheets.Add(variablesSheetName);
            variablesSheet = workbook.Worksheets[variablesSheetName];
        }

        variablesSheet.Cells["A1"].PutValue("varName");
        variablesSheet.Cells["B1"].PutValue("SampleValue");

        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.VariablesWorksheetName = variablesSheetName;
        designer.Process();

        string outputPath = "output.xlsx";
        workbook.Save(outputPath);

        Console.WriteLine($"Workbook saved to '{outputPath}' with variable 'varName' defined.");
    }
}