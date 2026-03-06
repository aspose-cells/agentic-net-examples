using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCustomSeparatorDemo
{
    class Program
    {
        static void Main()
        {
            // Path for the temporary CSV file
            string csvPath = "custom_separator.csv";

            // ---------- Create a workbook and fill sample data ----------
            Workbook workbook = new Workbook();                     // create workbook
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Alice");
            sheet.Cells["B3"].PutValue(25);

            // ---------- Save the workbook as a text file with a custom separator ----------
            TxtSaveOptions saveOptions = new TxtSaveOptions();      // create save options
            saveOptions.Separator = '*';                           // set custom separator
            workbook.Save(csvPath, saveOptions);                   // save using the rule

            // ---------- Load the text file using the same custom separator ----------
            TxtLoadOptions loadOptions = new TxtLoadOptions();      // create load options
            loadOptions.Separator = '*';                           // set the same separator
            Workbook loadedWorkbook = new Workbook(csvPath, loadOptions); // load using the rule

            // ---------- Verify loaded data ----------
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
            Console.WriteLine("A1: " + loadedSheet.Cells["A1"].StringValue);
            Console.WriteLine("B1: " + loadedSheet.Cells["B1"].StringValue);
            Console.WriteLine("A2: " + loadedSheet.Cells["A2"].StringValue);
            Console.WriteLine("B2: " + loadedSheet.Cells["B2"].IntValue);
            Console.WriteLine("A3: " + loadedSheet.Cells["A3"].StringValue);
            Console.WriteLine("B3: " + loadedSheet.Cells["B3"].IntValue);

            // Cleanup temporary file (optional)
            if (File.Exists(csvPath))
            {
                File.Delete(csvPath);
            }
        }
    }
}