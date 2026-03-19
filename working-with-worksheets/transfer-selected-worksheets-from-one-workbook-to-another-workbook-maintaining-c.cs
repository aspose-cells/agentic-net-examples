using System;
using Aspose.Cells;

namespace WorksheetTransferExample
{
    class Program
    {
        static void Main()
        {
            string sourceFile = "source.xlsx";
            string destinationFile = "destination.xlsx";

            Workbook sourceWorkbook = new Workbook(sourceFile);
            Workbook destinationWorkbook = new Workbook();
            destinationWorkbook.Worksheets.Clear();

            string[] sheetsToTransfer = { "Sheet1", "Data" };

            foreach (string sheetName in sheetsToTransfer)
            {
                Worksheet sourceSheet = sourceWorkbook.Worksheets[sheetName];
                if (sourceSheet == null)
                    continue; // Skip if the sheet does not exist in the source workbook

                Worksheet destSheet = destinationWorkbook.Worksheets.Add(sourceSheet.Name);

                CopyOptions copyOptions = new CopyOptions
                {
                    ReferToSheetWithSameName = true
                };

                sourceSheet.Copy(destSheet, copyOptions);
            }

            destinationWorkbook.Save(destinationFile, SaveFormat.Xlsx);
        }
    }
}