using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path for the sample XLS file
        string filePath = "sample.xls";

        // Create a new workbook and add sample data
        Workbook createWb = new Workbook();
        Worksheet ws = createWb.Worksheets[0];
        ws.Cells["A1"].PutValue("Name");
        ws.Cells["B1"].PutValue("Age");
        ws.Cells["A2"].PutValue("John");
        ws.Cells["B2"].PutValue(30);
        ws.Cells["A3"].PutValue("Jane");
        ws.Cells["B3"].PutValue(25);

        // Save the workbook as Excel 97‑2003 using XlsSaveOptions
        XlsSaveOptions saveOptions = new XlsSaveOptions();
        createWb.Save(filePath, saveOptions);

        // Load the XLS file with LoadOptions specifying the XLS format
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Excel97To2003);
        Workbook loadWb = new Workbook(filePath, loadOptions);

        // Access the first worksheet and display some cell values
        Worksheet loadedWs = loadWb.Worksheets[0];
        Console.WriteLine("A1: " + loadedWs.Cells["A1"].StringValue);
        Console.WriteLine("B2: " + loadedWs.Cells["B2"].StringValue);
    }
}