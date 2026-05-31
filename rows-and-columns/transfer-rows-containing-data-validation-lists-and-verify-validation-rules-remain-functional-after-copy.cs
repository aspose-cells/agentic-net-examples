using System;
using Aspose.Cells;

namespace AsposeCellsValidationCopyDemo
{
    class Program
    {
        static void Main()
        {
            // ---------- Create source workbook with a list validation ----------
            Workbook srcWorkbook = new Workbook();
            Worksheet srcSheet = srcWorkbook.Worksheets[0];

            // Populate source list values (A1:A3)
            srcSheet.Cells["A1"].PutValue("Apple");
            srcSheet.Cells["A2"].PutValue("Banana");
            srcSheet.Cells["A3"].PutValue("Cherry");

            // Add a List validation to cells B1:B3 that references A1:A3
            Validation srcValidation = srcSheet.Validations[srcSheet.Validations.Add()];
            srcValidation.Type = ValidationType.List;
            srcValidation.Formula1 = "A1:A3";
            // Apply validation to B1:B3 (row 0‑2, column 1)
            srcValidation.AddArea(CellArea.CreateCellArea(0, 1, 2, 1));

            // ---------- Create destination workbook ----------
            Workbook destWorkbook = new Workbook();
            Worksheet destSheet = destWorkbook.Worksheets[0];

            // ---------- Copy rows 0‑2 (first three rows) including validation ----------
            CopyOptions copyOptions = new CopyOptions(); // default options
            PasteOptions pasteOptions = new PasteOptions
            {
                PasteType = PasteType.Validation // ensure validations are copied
            };

            // CopyRows(sourceCells, sourceRowIndex, destinationRowIndex, rowNumber, copyOptions, pasteOptions)
            destSheet.Cells.CopyRows(
                srcSheet.Cells,   // source cells
                0,                // start at source row 0
                0,                // paste to destination row 0
                3,                // number of rows to copy (rows 0,1,2)
                copyOptions,
                pasteOptions);

            // ---------- Verify that validation was copied ----------
            Validation destValidation = destSheet.Validations.GetValidationInCell(0, 1); // cell B1
            if (destValidation != null)
            {
                Console.WriteLine("Validation successfully copied.");
                Console.WriteLine("Type: " + destValidation.Type);
                Console.WriteLine("Formula1: " + destValidation.Formula1);

                // Retrieve the first list value using GetListValue
                object firstValue = destValidation.GetListValue(0, 0);
                Console.WriteLine("First list value from copied validation: " + firstValue);
            }
            else
            {
                Console.WriteLine("Validation was not copied.");
            }

            // ---------- Save workbooks (optional) ----------
            srcWorkbook.Save("SourceWorkbook.xlsx", SaveFormat.Xlsx);
            destWorkbook.Save("DestinationWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}