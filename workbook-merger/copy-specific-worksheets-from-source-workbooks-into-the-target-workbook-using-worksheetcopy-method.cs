using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the source workbook that contains the worksheets to be copied
        Workbook sourceWorkbook = new Workbook("Source.xlsx");

        // Create a new (empty) target workbook
        Workbook targetWorkbook = new Workbook();

        // Remove the default worksheet that Aspose.Cells creates automatically
        targetWorkbook.Worksheets.Clear();

        // Define the names of the worksheets you want to copy from the source workbook
        string[] worksheetsToCopy = { "Sheet1", "Data" };

        foreach (string sheetName in worksheetsToCopy)
        {
            // Retrieve the source worksheet by name
            Worksheet sourceSheet = sourceWorkbook.Worksheets[sheetName];

            // Add a new blank worksheet to the target workbook with the same name
            Worksheet targetSheet = targetWorkbook.Worksheets.Add(sheetName);

            // Copy the contents and formatting from the source worksheet to the target worksheet
            targetSheet.Copy(sourceSheet);
        }

        // Save the target workbook with the copied worksheets
        targetWorkbook.Save("Target.xlsx");
    }
}