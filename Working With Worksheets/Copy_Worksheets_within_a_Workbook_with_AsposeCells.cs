using System;
using Aspose.Cells;

public class Program
{
    public static void Main()
    {
        Run();
    }

    public static void Run()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the default first worksheet and set its name and a sample value
        Worksheet original = workbook.Worksheets[0];
        original.Name = "Original";
        original.Cells["A1"].PutValue("Original Sheet Data");

        // Copy the original worksheet using AddCopy by sheet name
        int copiedIndex = workbook.Worksheets.AddCopy("Original");
        Worksheet copiedByName = workbook.Worksheets[copiedIndex];
        copiedByName.Name = "CopiedByName";
        copiedByName.Cells["A1"].PutValue("Copied Sheet Data");

        // Add a second worksheet with some data
        Worksheet second = workbook.Worksheets.Add("Second");
        second.Cells["B2"].PutValue(12345);

        // Copy a group of worksheets (original and second) to new sheets with specified names
        Worksheet[] sourceSheets = new Worksheet[] { original, second };
        string[] destNames = new string[] { "OriginalCopy", "SecondCopy" };
        workbook.Worksheets.AddCopy(sourceSheets, destNames);

        // Demonstrate Worksheet.Copy with CopyOptions (e.g., keep formula references to same‑named sheets)
        Worksheet destForCopy = workbook.Worksheets.Add("DestCopy");
        CopyOptions copyOptions = new CopyOptions();
        copyOptions.ReferToSheetWithSameName = true;
        destForCopy.Copy(original, copyOptions);

        // Save the workbook to a file
        workbook.Save("WorksheetCopyDemo.xlsx");
    }
}