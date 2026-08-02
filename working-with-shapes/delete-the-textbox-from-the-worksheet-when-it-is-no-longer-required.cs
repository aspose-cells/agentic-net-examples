// Title: Delete a TextBox shape from an Aspose.Cells worksheet (C#)
// Description: Shows how to create a workbook, add a TextBox, verify its existence, remove a specific TextBox (the last one in the collection) and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells delete textbox C# | remove textbox Aspose.Cells | TextBoxes.RemoveAt Aspose.Cells | C# delete shape Excel | Aspose.Cells shape management | worksheet TextBox removal | Aspose.Cells .NET delete shape
// Common Searches: How to delete a textbox in Excel with Aspose.Cells C# | Aspose.Cells remove TextBox shape programmatically | C# code to delete specific TextBox from worksheet | Delete all TextBoxes using Aspose.Cells .NET | Remove last TextBox from Aspose.Cells workbook
// Developer Intent: Remove one or more TextBox shapes from a worksheet using Aspose.Cells for .NET.
// Use Cases: Strip temporary annotation boxes before exporting a report. | Clean up user‑added textboxes during automated workbook processing. | Delete a textbox when a validation rule fails. | Purge all TextBoxes from a template workbook before reuse.
// AI Prompts: Write C# code that iterates through worksheet.TextBoxes and deletes each one with Aspose.Cells. | Show how to delete a TextBox by its index or by matching its text content in Aspose.Cells for .NET. | Explain safe removal of TextBox shapes without affecting other drawing objects in an Excel file using Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to create a workbook, add a TextBox, verify its existence, remove a specific TextBox (the last one in the collection) and save the file using Aspose.Cells for .NET.
public class DeleteTextboxDemo
{
    public static void Run()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            using (Workbook workbook = new Workbook())
            {
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a sample textbox (for demonstration purposes)
                worksheet.TextBoxes.Add(5, 5, 200, 50);
                worksheet.TextBoxes[0].Text = "Sample TextBox";

                // Delete the textbox if any exist
                if (worksheet.TextBoxes.Count > 0)
                {
                    // Remove the last textbox in the collection
                    int lastIndex = worksheet.TextBoxes.Count - 1;
                    worksheet.TextBoxes.RemoveAt(lastIndex);
                }

                // Save the workbook to a file
                string outputPath = "DeletedTextbox.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        DeleteTextboxDemo.Run();
    }
}
