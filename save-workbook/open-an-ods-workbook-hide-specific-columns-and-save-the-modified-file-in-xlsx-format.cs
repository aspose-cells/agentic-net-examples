// Title: C# – Hide Columns in an ODS Workbook and Convert to XLSX with Aspose.Cells
// Description: Load an ODS file using Aspose.Cells for .NET, hide selected columns (e.g., C‑E) on the first worksheet with the HideColumns method, and save the modified workbook as an XLSX file.
// Keywords: Aspose.Cells | C# | ODS to XLSX conversion | hide columns | HideColumns method | .NET spreadsheet API | column visibility | Excel export | OpenDocument Spreadsheet | programmatic column hiding
// Common Searches: Aspose.Cells hide columns in ODS | convert ODS to XLSX C# hide columns | C# hide column C in ODS using Aspose | how to hide multiple columns with Aspose.Cells | hide columns before saving ODS as Excel
// Developer Intent: Programmatically hide specific columns in an ODS file and export the result as an XLSX workbook.
// Use Cases: Remove confidential or layout columns from an ODS report before sharing it as an Excel file. | Create a template where certain columns are hidden for end‑users after conversion to XLSX. | Automate batch processing of ODS documents to hide unwanted columns and generate Excel‑compatible files.
// AI Prompts: Write C# code with Aspose.Cells that hides columns 5‑7 in an ODS workbook and saves it as XLSX. | Show an example that loads an ODS file, hides a range of columns based on their letters, and exports to XLSX with error handling. | Create a reusable method that accepts column letters to hide in an ODS file and returns the path of the saved XLSX file using Aspose.Cells.

using System;
using Aspose.Cells;

// Load an ODS file using Aspose.Cells for .NET, hide selected columns (e.g., C‑E) on the first worksheet with the HideColumns method, and save the modified workbook as an XLSX file.
class Program
{
    static void Main()
    {
        // Load the source ODS workbook
        Workbook workbook = new Workbook("input.ods");

        // Access the first worksheet (you can change the index as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Hide specific columns.
        // Example: hide columns C, D, and E (zero‑based indexes 2, 3, 4)
        // The HideColumns method takes the start column index and the number of columns to hide.
        worksheet.Cells.HideColumns(2, 3);

        // Save the modified workbook in XLSX format
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
