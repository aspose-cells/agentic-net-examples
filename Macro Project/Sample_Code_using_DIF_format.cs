using System;
using Aspose.Cells;
using Aspose.Cells.Loading;

namespace AsposeCellsDifDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Score");
            sheet.Cells["A2"].PutValue("Alice");
            sheet.Cells["B2"].PutValue(85);
            sheet.Cells["A3"].PutValue("Bob");
            sheet.Cells["B3"].PutValue(92);

            // Configure DIF save options
            DifSaveOptions difSaveOptions = new DifSaveOptions
            {
                ClearData = true,          // Make the workbook empty after saving
                CreateDirectory = true,    // Auto‑create directory if it does not exist
                RefreshChartCache = true   // Refresh chart cache (if any)
            };

            // Save the workbook in DIF format using the configured options
            string difPath = "SampleOutput.dif";
            workbook.Save(difPath, difSaveOptions);

            // Load the saved DIF file with default load options
            DifLoadOptions difLoadOptions = new DifLoadOptions();
            Workbook loadedWorkbook = new Workbook(difPath, difLoadOptions);

            // Access the loaded data (optional verification)
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
            Console.WriteLine("Loaded cell A2: " + loadedSheet.Cells["A2"].StringValue);
            Console.WriteLine("Loaded cell B2: " + loadedSheet.Cells["B2"].IntValue);

            // Save the loaded workbook to XLSX format for further verification
            loadedWorkbook.Save("Verification.xlsx", SaveFormat.Xlsx);
        }
    }
}