// Title: C# – Disable All Error Checks on a Named Worksheet with Aspose.Cells
// Description: A concise example that creates a new Workbook, renames its first worksheet, disables every ErrorCheckType using an ErrorCheckOption, applies the setting to the full used range, and returns the modified Worksheet for further processing or saving.
// Keywords: Aspose.Cells C# | disable worksheet error checks | ErrorCheckOptionCollection | ErrorCheckType false | set worksheet name Aspose.Cells | .NET Excel library example | apply error check to used range | GitHub Aspose.Cells sample | Excel error indicators off | programmatic worksheet configuration
// Common Searches: how to turn off all error checking in Aspose.Cells | Aspose.Cells disable error checks for entire sheet | C# set worksheet name and suppress error warnings | apply ErrorCheckOption to whole worksheet Aspose.Cells | sample code to disable Excel error indicators with Aspose
// Developer Intent: Create a method that accepts a worksheet name, disables every error‑check type on that sheet, and returns the configured Worksheet.
// Use Cases: Export data without Excel error triangles cluttering the view. | Provide a clean template for end‑users where validation warnings are unnecessary. | Generate multiple named sheets with uniform error‑check settings before populating data.
// AI Prompts: Write a C# function using Aspose.Cells that takes a worksheet name, disables all error checks for the entire used range, and returns the Worksheet object. | Show how to add an ErrorCheckOption to a worksheet, set each ErrorCheckType to false, and apply it to the full used area in Aspose.Cells for .NET. | Explain how to retrieve the parent Workbook from a returned Worksheet and save the file after disabling error checks.

using System;
using Aspose.Cells;

// A concise example that creates a new Workbook, renames its first worksheet, disables every ErrorCheckType using an ErrorCheckOption, applies the setting to the full used range, and returns the modified Worksheet for further processing or saving.
public static class WorksheetHelper
{
    /// <param name="sheetName">The name of the worksheet to create or rename.</param>
    /// <returns>The worksheet with all error checks disabled.</returns>
    public static Worksheet GetWorksheetWithAllErrorChecksDisabled(string sheetName)
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();

        // Use the first worksheet and set its name
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Name = sheetName;

        // Access the collection of error‑check options for the worksheet
        ErrorCheckOptionCollection options = worksheet.ErrorCheckOptions;

        // Add a new ErrorCheckOption that will hold our settings
        int optionIndex = options.Add();
        ErrorCheckOption option = options[optionIndex];

        // Disable every possible error‑check type
        foreach (ErrorCheckType type in Enum.GetValues(typeof(ErrorCheckType)))
        {
            option.SetErrorCheck(type, false);
        }

        // Apply the option to the whole used range of the worksheet
        int maxRow = worksheet.Cells.MaxRow;
        int maxCol = worksheet.Cells.MaxColumn;
        option.AddRange(CellArea.CreateCellArea(0, 0, maxRow, maxCol));

        // Return the modified worksheet
        return worksheet;
    }
}

public class Program
{
    public static void Main()
    {
        try
        {
            string sheetName = "MySheet";
            Worksheet ws = WorksheetHelper.GetWorksheetWithAllErrorChecksDisabled(sheetName);

            // Retrieve the workbook that owns the worksheet
            Workbook wb = ws.Workbook;

            // Save the workbook to verify the result
            string outputPath = "Output.xlsx";
            wb.Save(outputPath);
            Console.WriteLine($"Worksheet '{ws.Name}' created and saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
