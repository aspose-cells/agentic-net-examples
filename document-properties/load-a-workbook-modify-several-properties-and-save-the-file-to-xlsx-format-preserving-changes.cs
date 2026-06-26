using System;
using Aspose.Cells;

namespace AsposeCellsLoadModifySave
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file
            string sourcePath = "input.xlsx";

            // Load the workbook using the constructor that accepts a file path
            Workbook workbook = new Workbook(sourcePath);

            // ----- Modify workbook properties -----

            // 1. Change the default font for the entire workbook
            workbook.DefaultStyle.Font.Name = "Calibri";
            workbook.DefaultStyle.Font.Size = 11;

            // 2. Update built‑in document properties
            workbook.BuiltInDocumentProperties["Author"].Value = "Jane Doe";
            workbook.BuiltInDocumentProperties["Title"].Value = "Modified Workbook";

            // 3. Add a custom document property
            workbook.CustomDocumentProperties.Add("ReviewedBy", "John Smith");

            // 4. Rename the first worksheet
            Worksheet firstSheet = workbook.Worksheets[0];
            firstSheet.Name = "DataSheet";

            // ----- Save the modified workbook as XLSX -----
            // The Save method with file name and SaveFormat follows the provided rule.
            workbook.Save("output.xlsx", SaveFormat.Xlsx);

            // Optional: release resources
            workbook.Dispose();

            Console.WriteLine("Workbook loaded, modified, and saved successfully.");
        }
    }
}