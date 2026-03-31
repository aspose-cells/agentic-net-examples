using System;
using Aspose.Cells;
using Aspose.Cells.Loading;

public class DifReadWriteSample
{
    public static void Main()
    {
        // -------------------- Create workbook and add data --------------------
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Quantity");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(50);
        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["B3"].PutValue(30);

        // -------------------- Save as DIF using DifSaveOptions --------------------
        DifSaveOptions difSaveOptions = new DifSaveOptions
        {
            ClearData = true,          // clear workbook after saving
            CreateDirectory = true,    // create folder if it does not exist
            RefreshChartCache = true   // refresh chart cache (not used here but part of options)
        };
        workbook.Save("SampleData.dif", difSaveOptions);

        // -------------------- Load the DIF file using DifLoadOptions --------------------
        DifLoadOptions difLoadOptions = new DifLoadOptions();
        Workbook loadedWorkbook = new Workbook("SampleData.dif", difLoadOptions);
        Worksheet loadedSheet = loadedWorkbook.Worksheets[0];

        // -------------------- Read and display data from the loaded workbook --------------------
        Console.WriteLine("Data read from DIF file:");
        for (int row = 0; row <= 2; row++)
        {
            string product = loadedSheet.Cells[row, 0].StringValue;
            string quantity = loadedSheet.Cells[row, 1].StringValue;
            Console.WriteLine($"{product}\t{quantity}");
        }

        // -------------------- Save loaded workbook to another format for verification --------------------
        loadedWorkbook.Save("VerifiedOutput.xlsx", SaveFormat.Xlsx);
    }
}