using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        string templatePath = "template.xltx";
        string jsonPath = "data.json";
        string outputPath = "result.xlsx";

        Workbook workbook = new Workbook(templatePath);
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        string jsonData = File.ReadAllText(jsonPath);
        designer.SetJsonDataSource("ds", jsonData);
        designer.Process();
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}