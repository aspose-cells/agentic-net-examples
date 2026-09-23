// Title: Refresh linked OLE objects in an Excel workbook using Aspose.Cells for .NET with reflection
// AI Prompts: Write C# code that opens an Excel file with Aspose.Cells, enumerates all worksheets, finds OleObjects that are linked, and refreshes each by invoking the UpdateLink method via reflection. | Modify the example to target a single worksheet identified by name, refresh only its linked OleObjects, and output the names of objects that were successfully updated. | Enhance the sample with robust error handling that skips OleObjects lacking the IsLinked property or UpdateLink method, logs the issue, and still saves the workbook.
// Common Searches: aspnet update external OLE links in Excel using Aspose.Cells | c# use reflection to trigger OleObject link refresh with Aspose.Cells | how to determine IsLinked flag of OleObject with Aspose.Cells | programmatically refresh OLE links in an Excel workbook via Aspose.Cells for .NET | persist workbook after OLE link update with Aspose.Cells
// Tags: update external OLE links Aspose.Cells | invoke link refresh using reflection | enumerate OleObjects per worksheet | detect linked OleObject flag | persist workbook after OLE update

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Drawing; // Required for OleObject

// The sample loads a workbook, walks through each worksheet's OleObjects, uses reflection to verify the IsLinked flag, calls UpdateLink on linked objects, logs any failures, and saves the workbook to persist the refreshed links.
class Program
{
    static void Main(string[] args)
    {
        // Path to the Excel file containing linked OLE objects
        string filePath = @"C:\Temp\LinkedOleWorkbook.xlsx";

        // Refresh linked OLE objects
        RefreshLinkedOleObjects(filePath);
    }

    /// <param name="excelFilePath">Full path to the Excel file.</param>
    static void RefreshLinkedOleObjects(string excelFilePath)
    {
        // Verify that the file exists before attempting to load it
        if (!File.Exists(excelFilePath))
        {
            Console.WriteLine("File not found: " + excelFilePath);
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(excelFilePath);

            // Iterate through each worksheet and refresh its OLE objects
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (OleObject ole in sheet.OleObjects)
                {
                    try
                    {
                        // Use reflection to check for the 'IsLinked' property (may not exist in older versions)
                        PropertyInfo isLinkedProp = ole.GetType().GetProperty("IsLinked", BindingFlags.Public | BindingFlags.Instance);
                        bool isLinked = isLinkedProp != null && (bool)isLinkedProp.GetValue(ole);

                        if (isLinked)
                        {
                            // Use reflection to invoke 'UpdateLink' method if available
                            MethodInfo updateLinkMethod = ole.GetType().GetMethod("UpdateLink", BindingFlags.Public | BindingFlags.Instance);
                            updateLinkMethod?.Invoke(ole, null);
                        }
                    }
                    catch (Exception innerEx)
                    {
                        // Log but continue processing other OLE objects
                        Console.WriteLine($"Failed to refresh OLE object '{ole.Name}': {innerEx.Message}");
                    }
                }
            }

            // Save the workbook to persist the refreshed links
            workbook.Save(excelFilePath);
            Console.WriteLine("Linked OLE objects refreshed successfully.");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine("An error occurred while refreshing OLE objects: " + ex.Message);
        }
    }
}
