using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data
        worksheet.Cells["A1"].PutValue("Name");
        worksheet.Cells["B1"].PutValue("Score");
        worksheet.Cells["A2"].PutValue("Alice");
        worksheet.Cells["B2"].PutValue(85);
        worksheet.Cells["A3"].PutValue("Bob");
        worksheet.Cells["B3"].PutValue(92);
        worksheet.Cells["A4"].PutValue("Charlie");
        worksheet.Cells["B4"].PutValue(78);

        // Configure DIF save options
        DifSaveOptions difOptions = new DifSaveOptions
        {
            ClearData = true,
            CreateDirectory = true,
            RefreshChartCache = true
        };

        // Define temporary DIF file path
        string difFilePath = Path.Combine(Path.GetTempPath(), "sample.dif");

        // Save the workbook in DIF format using the options
        workbook.Save(difFilePath, difOptions);

        // Read the generated DIF file and output its content to the console
        string difContent = File.ReadAllText(difFilePath);
        Console.WriteLine(difContent);
    }
}