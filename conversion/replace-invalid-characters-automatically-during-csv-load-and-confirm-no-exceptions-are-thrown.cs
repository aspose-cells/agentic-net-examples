// Title: C# – Load CSV with invalid Excel characters using Aspose.Cells (disable restriction checks)
// Description: Demonstrates how to create a CSV containing null (\u0000) and unit‑separator (\u001F) characters, configure TxtLoadOptions with CheckExcelRestriction = false, load the file into a Workbook without exceptions, display the cell values, and save the result as XLSX.
// Keywords: Aspose.Cells CSV load invalid characters | CheckExcelRestriction false | TxtLoadOptions CSV import C# | ignore Excel restrictions Aspose | CSV to XLSX conversion Aspose.Cells | C# load CSV with control characters | Aspose.Cells data preprocessing
// Common Searches: load CSV containing null characters with Aspose.Cells | Aspose.Cells ignore Excel restriction when importing CSV | replace invalid characters automatically during CSV load Aspose | Aspose.Cells TxtLoadOptions CheckExcelRestriction example | C# load CSV with control characters without error
// Developer Intent: Load a CSV file that includes characters prohibited by Excel while preventing any import or save exceptions.
// Use Cases: Import legacy CSV logs that contain control characters and convert them to XLSX for reporting. | Migrate data from older systems where CSV exports may include null or non‑printable bytes. | Create a preprocessing routine that safely reads CSV data into a workbook by bypassing Excel's character restrictions.
// AI Prompts: Show how to set TxtLoadOptions.CheckExcelRestriction to false for CSV loading in Aspose.Cells (C#). | Write C# code that loads a CSV with null and unit‑separator characters into a Workbook without errors using Aspose.Cells. | Explain the effect of CheckExcelRestriction on CSV import and how to handle invalid characters in Aspose.Cells.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// Demonstrates how to create a CSV containing null (\u0000) and unit‑separator (\u001F) characters, configure TxtLoadOptions with CheckExcelRestriction = false, load the file into a Workbook without exceptions, display the cell values, and save the result as XLSX.
class CsvLoadReplaceInvalidDemo
{
    static void Main()
    {
        // Prepare a temporary CSV file that contains characters invalid for Excel cells
        string tempCsvPath = "temp_invalid.csv";
        string csvContent = "Name,Comment\nJohn,\"Hello\u0000World\"\nJane,\"Good\u001FDay\"";
        File.WriteAllText(tempCsvPath, csvContent, Encoding.UTF8);

        // Configure load options to ignore Excel restrictions (invalid characters will be accepted)
        TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv)
        {
            Encoding = Encoding.UTF8,
            // Disabling restriction checking prevents exceptions caused by invalid characters
            // This property is inherited from LoadOptions
            CheckExcelRestriction = false
        };

        // Load the CSV file into a workbook using the configured options
        Workbook workbook = new Workbook(tempCsvPath, loadOptions);

        // Access the first worksheet and output the imported values to verify successful load
        Worksheet sheet = workbook.Worksheets[0];
        Console.WriteLine("A2 (Name): " + sheet.Cells["A2"].StringValue);
        Console.WriteLine("B2 (Comment): " + sheet.Cells["B2"].StringValue);
        Console.WriteLine("A3 (Name): " + sheet.Cells["A3"].StringValue);
        Console.WriteLine("B3 (Comment): " + sheet.Cells["B3"].StringValue);

        // Save the workbook to confirm that no exception occurs during the save operation
        workbook.Save("output.xlsx", SaveFormat.Xlsx);

        // Clean up the temporary CSV file
        File.Delete(tempCsvPath);
    }
}
