using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        using (Workbook workbook = new Workbook())
        {
            Worksheet sheet = workbook.Worksheets[0];

            sheet.Cells["A1"].PutValue("Sample Text");
            sheet.Cells["B2"].PutValue(42);

            TxtSaveOptions saveOptions = new TxtSaveOptions
            {
                ClearData = true,
                Separator = ','
            };
            workbook.Save("output.txt", saveOptions);

            SheetRender renderer = new SheetRender(sheet, new ImageOrPrintOptions());
            renderer.ToImage(0, "Sheet0.png");
        }
    }
}