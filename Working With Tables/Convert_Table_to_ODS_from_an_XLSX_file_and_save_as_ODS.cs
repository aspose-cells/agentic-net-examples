using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Ods;

namespace AsposeCellsExamples
{
    public class TableToOdsConversion
    {
        public static void Run()
        {
            // Path to the source XLSX file containing a table (ListObject)
            string sourcePath = "input.xlsx";

            // Path for the resulting ODS file
            string destPath = "output.ods";

            // Load the workbook from the XLSX file
            Workbook workbook = new Workbook(sourcePath);

            // Iterate through all worksheets and convert each table to a normal range
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // ListObjects collection holds all tables in the worksheet
                for (int i = 0; i < sheet.ListObjects.Count; i++)
                {
                    ListObject table = sheet.ListObjects[i];
                    // Convert the table to a regular cell range
                    table.ConvertToRange();
                }
            }

            // Create ODS save options (optional: set generator type)
            OdsSaveOptions saveOptions = new OdsSaveOptions
            {
                GeneratorType = OdsGeneratorType.LibreOffice
            };

            // Save the modified workbook as ODS
            workbook.Save(destPath, saveOptions);

            Console.WriteLine($"Conversion completed: '{sourcePath}' -> '{destPath}'");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            TableToOdsConversion.Run();
        }
    }
}