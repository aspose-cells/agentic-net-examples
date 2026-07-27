// Title: Save a Modified Workbook to XLSX with Aspose.Cells Workbook.Save (C#)
// Description: Demonstrates how to create a Workbook, change cell values (string, numeric, date) and persist the changes by calling Workbook.Save with SaveFormat.Xlsx and the default save options.
// Keywords: Aspose.Cells C# save workbook | Workbook.Save default options | SaveFormat.Xlsx example | modify cell values Aspose.Cells | export Excel to XLSX C# | Aspose.Cells Save method | C# Excel file generation
// Common Searches: Aspose.Cells save workbook as XLSX C# | Workbook.Save overload default options example | how to write changes to Excel file using Aspose.Cells | C# code to modify cells and save as XLSX | Aspose.Cells default save settings
// Developer Intent: Persist edited cell data by saving the workbook to an XLSX file using the default save configuration.
// Use Cases: Create a new workbook, populate cells with different data types, and generate an XLSX report. | Update an existing workbook's content and overwrite the original file without custom save parameters. | Automate simple data entry tasks where values are written to specific cells and the file is saved in standard Excel format.
// AI Prompts: Generate C# code that opens an Excel file with Aspose.Cells, updates several cells, and saves it as XLSX using Workbook.Save with default options. | Explain the steps to use Workbook.Save and SaveFormat.Xlsx to export a modified workbook after changing string, numeric, and date cells. | Show how to overwrite an existing XLSX file after editing cell values with Aspose.Cells in C#.

using System;
using Aspose.Cells;

namespace AsposeCellsSaveExample
{
    // Demonstrates how to create a Workbook, change cell values (string, numeric, date) and persist the changes by calling Workbook.Save with SaveFormat.Xlsx and the default save options.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses Workbook() constructor rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Modify some cell values
            sheet.Cells["A1"].PutValue("First");
            sheet.Cells["B2"].PutValue(123);
            sheet.Cells["C3"].PutValue(DateTime.Now);

            // Save the workbook back to XLSX format using default options
            // The Save(string, SaveFormat) overload follows the provided rule.
            workbook.Save("ModifiedWorkbook.xlsx", SaveFormat.Xlsx);

            // Optional: inform the user
            Console.WriteLine("Workbook saved as ModifiedWorkbook.xlsx");
        }
    }
}
