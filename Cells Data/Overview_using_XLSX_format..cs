using System;
using System.Threading.Tasks;
using Aspose.Cells;

class Program
{
    static async Task Main(string[] args)
    {
        // Initialize a new workbook (create)
        Workbook workbook = new Workbook();

        // Access the default worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Overview";

        // Populate some sample data
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Amount");
        sheet.Cells["A2"].PutValue("Revenue");
        sheet.Cells["B2"].PutValue(125000);
        sheet.Cells["A3"].PutValue("Cost");
        sheet.Cells["B3"].PutValue(73000);
        sheet.Cells["A4"].PutValue("Profit");
        sheet.Cells["B4"].PutValue(52000);

        // Define the output XLSX file path (save)
        string outputPath = "Overview.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);

        // Output confirmation
        Console.WriteLine($"Workbook saved to {outputPath}");
        await Task.CompletedTask;
    }
}