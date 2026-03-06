using System;
using Aspose.Cells;

public class OpenWorkbookDemo
{
    public static void Main()
    {
        // Path to the source Excel file (can be .xls, .xlsx, .xlsb, .csv, etc.)
        string sourcePath = "sample.xlsx";

        // Load the workbook using the constructor that auto‑detects the format
        Workbook workbook = new Workbook(sourcePath);

        // Display the detected file format
        Console.WriteLine($"Loaded workbook format: {workbook.FileFormat}");

        // Example of using LoadOptions (e.g., for password‑protected files)
        // LoadOptions loadOptions = new LoadOptions();
        // loadOptions.Password = "yourPassword";
        // Workbook workbookWithOptions = new Workbook(sourcePath, loadOptions);

        // Modify the workbook (add a simple value)
        workbook.Worksheets[0].Cells["A1"].PutValue("Opened successfully");

        // Save the workbook to a different format (PDF in this case)
        string outputPath = "output.pdf";
        workbook.Save(outputPath, SaveFormat.Pdf);

        Console.WriteLine($"Workbook saved as {outputPath}");
    }
}