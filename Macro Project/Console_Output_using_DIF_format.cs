using System;
using Aspose.Cells;
using Aspose.Cells.Loading;

class Program
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Sample Text");
        worksheet.Cells["A2"].PutValue(12345);
        worksheet.Cells["B1"].PutValue(DateTime.Now);

        // Configure DIF save options
        DifSaveOptions difSaveOptions = new DifSaveOptions
        {
            ClearData = true,          // make workbook empty after saving
            CreateDirectory = true,    // create directory if it does not exist
            RefreshChartCache = true   // refresh chart cache (if any)
        };

        // Save the workbook in DIF format
        string difFilePath = "SampleOutput.dif";
        workbook.Save(difFilePath, difSaveOptions);
        Console.WriteLine($"Workbook saved to DIF format at: {difFilePath}");

        // Load the saved DIF file using DifLoadOptions
        DifLoadOptions difLoadOptions = new DifLoadOptions();
        Workbook loadedWorkbook = new Workbook(difFilePath, difLoadOptions);
        Worksheet loadedWorksheet = loadedWorkbook.Worksheets[0];

        // Retrieve and display loaded values
        string loadedA1 = loadedWorksheet.Cells["A1"].StringValue;
        string loadedA2 = loadedWorksheet.Cells["A2"].StringValue;
        string loadedB1 = loadedWorksheet.Cells["B1"].StringValue;

        Console.WriteLine($"Loaded values -> A1: {loadedA1}, A2: {loadedA2}, B1: {loadedB1}");
    }
}