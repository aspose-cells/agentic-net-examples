// Title: Aspose.Cells for .NET – Enable UpdateReference before Deleting Blank Rows & Columns
// Description: C# sample that configures DeleteBlankOptions (sets UpdateReference = true and treats empty strings as blanks) and then removes empty rows and columns from a worksheet while automatically adjusting formulas and references throughout the workbook.
// Keywords: Aspose.Cells | C# DeleteBlankRows | DeleteBlankColumns | DeleteBlankOptions | UpdateReference property | empty string as blank | preserve Excel formulas | Excel automation .NET | GitHub Aspose.Cells example | US region | Europe region
// Common Searches: Aspose.Cells set UpdateReference true | Delete blank rows keep formulas | Treat empty strings as blanks Aspose.Cells | Remove blank columns with reference update .NET | Aspose.Cells DeleteBlankOptions GitHub sample
// Developer Intent: Configure DeleteBlankOptions with UpdateReference enabled and use it to delete blank rows and columns.
// Use Cases: Clean imported data by removing empty rows while retaining dependent formulas. | Eliminate blank columns that contain empty strings without breaking chart data sources. | Apply identical blank‑row/column removal across all worksheets in a large workbook. | Prepare a workbook for distribution by stripping unused rows and columns automatically.
// AI Prompts: Generate C# code that iterates through every worksheet and deletes blank rows and columns using DeleteBlankOptions with UpdateReference set to true. | Explain how the UpdateReference flag influences formulas when blank rows are removed in Aspose.Cells. | Provide a step‑by‑step guide to treat empty‑string cells as blanks and keep drawings intact during deletion. | Show how to log which rows and columns were removed after calling DeleteBlankRows/DeleteBlankColumns.

using System;
using Aspose.Cells;

// C# sample that configures DeleteBlankOptions (sets UpdateReference = true and treats empty strings as blanks) and then removes empty rows and columns from a worksheet while automatically adjusting formulas and references throughout the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate the worksheet with sample data that includes blank rows and columns
        cells["A1"].PutValue("Header");
        cells["A2"].PutValue("");               // Blank row (empty string)
        cells["A3"].PutValue("Data");
        cells["B1"].PutValue("");               // Blank column (empty string)
        cells["C1"].PutValue("Another Header");
        cells["C2"].PutValue("More Data");

        // Configure DeleteBlankOptions (inherits DeleteOptions) and ensure UpdateReference is true
        DeleteBlankOptions deleteOptions = new DeleteBlankOptions
        {
            UpdateReference = true,          // Important: update references in other worksheets
            EmptyStringAsBlank = true,       // Treat empty strings as blanks
            DrawingsAsBlank = true           // Default behavior for drawings
        };

        // Delete blank rows using the configured options
        worksheet.Cells.DeleteBlankRows(deleteOptions);

        // Delete blank columns using the same options
        worksheet.Cells.DeleteBlankColumns(deleteOptions);

        // Save the modified workbook
        workbook.Save("ProcessedWorkbook.xlsx", SaveFormat.Xlsx);
    }
}
