using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class CombineAndSaveWorkbook
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            // Create the source workbook and add sample data
            Workbook sourceWorkbook = new Workbook();
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            sourceSheet.Name = "Source";
            sourceSheet.Cells["A1"].PutValue("Source Data");
            sourceSheet.Cells["B2"].PutValue(123);

            // Create the destination workbook and add its own data
            Workbook destWorkbook = new Workbook();
            Worksheet destSheet = destWorkbook.Worksheets[0];
            destSheet.Name = "Destination";
            destSheet.Cells["A1"].PutValue("Destination Data");
            destSheet.Cells["C3"].PutValue(456);

            // Combine the source workbook into the destination workbook
            destWorkbook.Combine(sourceWorkbook);

            // Save the combined workbook as an XLSX file (Excel compatible)
            destWorkbook.Save("CombinedWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}