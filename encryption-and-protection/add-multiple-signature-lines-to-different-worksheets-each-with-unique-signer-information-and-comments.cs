// Title: Create unique signature lines with signer details on multiple Excel worksheets using Aspose.Cells for .NET
// Common Searches: how to add a different signature textbox to each sheet in an Excel workbook with Aspose.Cells C# | Aspose.Cells C# example for inserting signer name, title, email, and comments on separate worksheets | C# code to place a free‑floating text box as a signature line on multiple Excel sheets | add custom signer information to finance, HR, and IT worksheets using Aspose.Cells | Aspose.Cells .NET create multiple signature lines with comments per worksheet
// Tags: add signature textbox per worksheet Aspose.Cells | custom signer details Excel shape C# | freefloating shape placement Aspose.Cells | multiple worksheet signature lines Aspose.Cells | Aspose.Cells add text box shape example

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new workbook, clears the default sheet, adds three worksheets (Finance, HR, IT) and inserts a signature line into each by adding a free‑floating text box shape that displays the signer’s name, title, email, and comments. Each shape is configured with PlacementType.FreeFloating, and the workbook is saved as MultipleSignatureLines.xlsx with error handling for both workbook creation and signature insertion.
class AddMultipleSignatureLines
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Remove the default sheet (optional) and add custom worksheets
            workbook.Worksheets.Clear();

            // Worksheet 1 with its own signature line
            Worksheet sheet1 = workbook.Worksheets.Add("Finance");
            AddSignatureLine(sheet1, "Alice Johnson", "Chief Financial Officer", "alice.johnson@example.com", "Approved for Q1 financials.");

            // Worksheet 2 with its own signature line
            Worksheet sheet2 = workbook.Worksheets.Add("HR");
            AddSignatureLine(sheet2, "Bob Smith", "HR Manager", "bob.smith@example.com", "Reviewed employee records.");

            // Worksheet 3 with its own signature line
            Worksheet sheet3 = workbook.Worksheets.Add("IT");
            AddSignatureLine(sheet3, "Carol Lee", "IT Director", "carol.lee@example.com", "Validated system architecture.");

            // Save the workbook to a file
            workbook.Save("MultipleSignatureLines.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    /// <param name="worksheet">Target worksheet.</param>
    /// <param name="signerName">Name of the signer.</param>
    /// <param name="signerTitle">Title of the signer.</param>
    /// <param name="signerEmail">Email of the signer.</param>
    /// <param name="comments">Comments to display with the signature line.</param>
    private static void AddSignatureLine(Worksheet worksheet, string signerName, string signerTitle, string signerEmail, string comments)
    {
        try
        {
            // Create a text box shape to simulate a signature line
            // Parameters: upper left row, upper left column, row offset, column offset, height (pixels), width (pixels)
            Shape signatureShape = worksheet.Shapes.AddTextBox(5, 1, 0, 0, 100, 300);

            // Set the displayed text with signer details
            signatureShape.Text = $"Signed by: {signerName}\nTitle: {signerTitle}\nEmail: {signerEmail}\nComments: {comments}";

            // Optional: set the shape's placement to move and size with cells
            signatureShape.Placement = PlacementType.FreeFloating;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to add signature line on sheet '{worksheet.Name}': {ex.Message}");
        }
    }
}
