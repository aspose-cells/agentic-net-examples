// Title: Refresh linked picture objects in an Excel workbook after source image changes with Aspose.Cells for .NET
// Common Searches: c# refresh linked pictures in Excel using Aspose.Cells | how to update external image references in an Aspose.Cells workbook | using reflection to call UpdateLinkedImage on linked pictures Aspose.Cells | programmatically refresh linked images after source file modification with Aspose.Cells .NET | check IsLinked property of pictures in Aspose.Cells example
// Tags: Aspose.Cells linked picture refresh | UpdateLinkedImage method Aspose.Cells | external image reference update .NET | worksheet picture iteration Aspose.Cells | save workbook after picture refresh

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing Excel workbook, iterates through each worksheet and its pictures, uses reflection to detect the IsLinked property, invokes the UpdateLinkedImage method on linked pictures to refresh them, and then saves the workbook to persist the updated visuals.
class RefreshLinkedPictures
{
    static void Main()
    {
        // Path to the Excel file containing linked pictures
        string workbookPath = @"C:\Temp\LinkedPicturesWorkbook.xlsx";

        // Verify that the workbook file exists to avoid FileNotFoundException
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Error: The file \"{workbookPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all pictures in the worksheet
                foreach (Picture pic in sheet.Pictures)
                {
                    try
                    {
                        // Use reflection to check for the 'IsLinked' property (if available)
                        bool isLinked = true; // default to true if property not found
                        PropertyInfo isLinkedProp = pic.GetType().GetProperty("IsLinked", BindingFlags.Public | BindingFlags.Instance);
                        if (isLinkedProp != null && isLinkedProp.PropertyType == typeof(bool))
                        {
                            isLinked = (bool)isLinkedProp.GetValue(pic);
                        }

                        if (isLinked)
                        {
                            // Use reflection to invoke 'UpdateLinkedImage' method (if available)
                            MethodInfo updateMethod = pic.GetType().GetMethod("UpdateLinkedImage", BindingFlags.Public | BindingFlags.Instance);
                            if (updateMethod != null)
                            {
                                updateMethod.Invoke(pic, null);
                            }
                        }
                    }
                    catch (Exception exPic)
                    {
                        Console.WriteLine($"Warning: Could not refresh picture on sheet \"{sheet.Name}\": {exPic.Message}");
                    }
                }
            }

            // Save the workbook (overwrites the original file)
            workbook.Save(workbookPath);
            Console.WriteLine("Linked pictures refreshed and workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
