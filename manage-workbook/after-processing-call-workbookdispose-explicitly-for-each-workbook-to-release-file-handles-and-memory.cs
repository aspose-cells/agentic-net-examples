using System;
using Aspose.Cells;

namespace AsposeCellsDisposeDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // ------------------------------
            // Create a new workbook, add data, save and dispose it.
            // ------------------------------
            Workbook workbook = new Workbook();                     // create workbook (rule: Workbook())
            Worksheet sheet = workbook.Worksheets[0];              // access default worksheet
            sheet.Cells["A1"].PutValue("Hello, Aspose.Cells!");    // add sample data
            sheet.Cells["B2"].PutValue(DateTime.Now);              // add another sample value

            // Save the workbook to a file (rule: Save(string))
            string createdFile = "CreatedWorkbook.xlsx";
            workbook.Save(createdFile);                            // save workbook

            // Explicitly release resources
            workbook.Dispose();                                    // dispose workbook

            // ------------------------------
            // Load an existing workbook, modify it, save and dispose it.
            // ------------------------------
            // Ensure the file exists; for demo purposes we reuse the file we just created.
            Workbook loadedWorkbook = new Workbook(createdFile);   // load workbook (rule: Workbook(string))
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
            loadedSheet.Cells["C3"].PutValue("Added after load"); // modify workbook

            // Save the modified workbook (rule: Save(string, SaveFormat))
            string modifiedFile = "ModifiedWorkbook.xlsx";
            loadedWorkbook.Save(modifiedFile, SaveFormat.Xlsx);   // save with format

            // Explicitly release resources
            loadedWorkbook.Dispose();                              // dispose loaded workbook
        }
    }
}