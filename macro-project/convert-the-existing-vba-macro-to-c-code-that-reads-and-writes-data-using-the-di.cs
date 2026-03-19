using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Loading;
using Aspose.Cells.Saving;

namespace AsposeCellsDifExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Paths to the source and destination DIF files
            string sourceDifPath = "SourceData.dif";
            string destDifPath = "ModifiedData.dif";

            // Ensure the source DIF file exists; create a simple one if it doesn't
            if (!File.Exists(sourceDifPath))
            {
                Workbook tempWb = new Workbook();
                tempWb.Worksheets[0].Cells["A1"].PutValue("Initial Value");
                tempWb.Save(sourceDifPath, new DifSaveOptions());
            }

            // Load the workbook from the DIF file using DIF load options
            DifLoadOptions loadOptions = new DifLoadOptions();
            Workbook workbook = new Workbook(sourceDifPath, loadOptions);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Read and display the original value from cell A1
            string originalValue = sheet.Cells["A1"].StringValue;
            Console.WriteLine($"Original A1 value: {originalValue}");

            // Modify the value in cell A1
            sheet.Cells["A1"].PutValue("Updated via Aspose.Cells");

            // Save the workbook to a new DIF file using DIF save options
            DifSaveOptions saveOptions = new DifSaveOptions
            {
                ClearData = true,
                CreateDirectory = true,
                RefreshChartCache = true
            };
            workbook.Save(destDifPath, saveOptions);

            Console.WriteLine($"Workbook saved to DIF format at: {destDifPath}");
        }
    }
}