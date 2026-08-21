// Title: Delete a TextBox from an Aspose.Cells Worksheet (C#/.NET)
// Description: Shows how to create a workbook, add a temporary TextBox, verify the TextBoxes collection, remove the first TextBox with worksheet.TextBoxes.RemoveAt(0), and save the file, demonstrating the proper way to delete a TextBox when it is no longer needed.
// Keywords: Aspose.Cells delete textbox | C# TextBox removal Aspose.Cells | worksheet.TextBoxes.RemoveAt example | remove temporary textbox .NET | Aspose.Cells shape deletion | Aspose.Cells .NET API
// Common Searches: Aspose.Cells C# delete textbox | Remove TextBox from Excel sheet using Aspose.Cells | How to clear all TextBoxes in a worksheet with Aspose.Cells | Delete specific TextBox by index Aspose.Cells .NET | Programmatically remove shapes from workbook Aspose.Cells
// Developer Intent: Remove a TextBox that was added to a worksheet once it is no longer required.
// Use Cases: Clean up temporary annotation boxes before exporting the workbook to avoid unwanted UI elements. | Strip user‑added TextBoxes during batch processing of multiple reports. | Ensure generated Excel files contain no stray TextBoxes when converting to PDF or other formats.
// AI Prompts: Provide C# code that deletes all TextBoxes from a worksheet using Aspose.Cells for .NET. | Show an example that finds a TextBox with specific text and removes it from an Aspose.Cells worksheet. | Explain how to iterate through worksheet.TextBoxes safely and delete each item.

using System;
using Aspose.Cells;

// Shows how to create a workbook, add a temporary TextBox, verify the TextBoxes collection, remove the first TextBox with worksheet.TextBoxes.RemoveAt(0), and save the file, demonstrating the proper way to delete a TextBox when it is no longer needed.
public class DeleteTextboxDemo
{
    public static void Main(string[] args)
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Create a new workbook and get the first worksheet
        using (Workbook workbook = new Workbook())
        {
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a textbox to the worksheet (for demonstration)
            worksheet.TextBoxes.Add(10, 10, 200, 50);
            worksheet.TextBoxes[0].Text = "Temporary TextBox";

            // Remove the textbox when it is no longer required
            if (worksheet.TextBoxes.Count > 0)
            {
                // Remove the textbox at index 0
                worksheet.TextBoxes.RemoveAt(0);
            }

            // Save the workbook to verify the textbox has been removed
            workbook.Save("DeletedTextboxDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}
