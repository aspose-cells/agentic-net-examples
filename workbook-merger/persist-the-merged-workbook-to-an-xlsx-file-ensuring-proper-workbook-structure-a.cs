using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create the first workbook and add some data
        Workbook firstWorkbook = new Workbook();
        Worksheet firstSheet = firstWorkbook.Worksheets[0];
        firstSheet.Name = "FirstSheet";
        firstSheet.Cells["A1"].PutValue("First Workbook");
        firstSheet.Cells["A2"].PutValue(123);
        // Merge cells A1:C1 in the first sheet
        firstSheet.Cells.Merge(0, 0, 1, 2);

        // Create the second workbook and add some data
        Workbook secondWorkbook = new Workbook();
        Worksheet secondSheet = secondWorkbook.Worksheets[0];
        secondSheet.Name = "SecondSheet";
        secondSheet.Cells["A1"].PutValue("Second Workbook");
        secondSheet.Cells["A2"].PutValue(456);
        // Merge cells A1:B1 in the second sheet
        secondSheet.Cells.Merge(0, 0, 1, 1);

        // Combine the second workbook into the first workbook
        firstWorkbook.Combine(secondWorkbook);

        // Configure save options for XLSX with merged‑area validation
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
        {
            ValidateMergedAreas = true, // ensure merged cells are valid before saving
            MergeAreas = true           // merge conditional formatting/validation areas
        };

        // Persist the merged workbook to an XLSX file
        firstWorkbook.Save("MergedWorkbook.xlsx", saveOptions);
    }
}