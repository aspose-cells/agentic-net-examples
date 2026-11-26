// Title: Set OleObject.IsHidden to true for icon‑displayed OLE objects in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# hide OLE objects that are displayed as icons | How to set OleObject.IsHidden property in a .NET Excel file | Programmatically hide background OLE objects in an Aspose.Cells workbook | C# code to hide embedded OLE icons in Excel using Aspose.Cells | Make OLE objects invisible in an Excel worksheet with Aspose.Cells for .NET
// Tags: Aspose.Cells OleObject IsHidden | C# hide OLE icon objects Excel | Aspose.Cells set OLE object visibility | Excel workbook hide embedded OLE objects | OleObject DisplayAsIcon handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace OleObjectVisibilityDemo
{
    // The example loads or creates an Excel workbook, adds sample OLE objects (one shown as an icon, one displayed directly), iterates through all OleObject instances on the first worksheet, sets IsHidden to true for those with DisplayAsIcon enabled, ensures other objects stay visible, and saves the workbook with the updated OLE object visibility.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Prepare a workbook.
                // -----------------------------------------------------------------
                Workbook workbook;
                string inputPath = "InputWorkbook.xlsx";

                if (File.Exists(inputPath))
                {
                    // Load an existing workbook if present.
                    workbook = new Workbook(inputPath);
                    Console.WriteLine($"Loaded workbook from \"{inputPath}\".");
                }
                else
                {
                    // Otherwise create a new workbook with a single worksheet.
                    workbook = new Workbook();
                    workbook.Worksheets[0].Name = "Sheet1";
                    Console.WriteLine("Created a new workbook.");
                }

                Worksheet sheet = workbook.Worksheets[0];

                // -----------------------------------------------------------------
                // 2. Add sample OLE objects (if the worksheet does not already contain any).
                // -----------------------------------------------------------------
                if (sheet.OleObjects.Count == 0)
                {
                    // Prepare dummy data for the OLE objects.
                    // In a real scenario replace these with actual file bytes.
                    byte[] dummyImage = CreateDummyImageBytes();
                    byte[] dummyExcel = new Workbook().SaveToStream().ToArray();

                    // Add an OLE object that displays as an icon.
                    int oleIndex1 = sheet.OleObjects.Add(5, 2, 150, 150, dummyImage);
                    OleObject ole1 = sheet.OleObjects[oleIndex1];
                    ole1.DisplayAsIcon = true;
                    ole1.Label = "Embedded Excel File";
                    ole1.ObjectData = dummyExcel;
                    ole1.ProgID = "Excel.Sheet.12";

                    // Add a second OLE object that shows the image directly.
                    int oleIndex2 = sheet.OleObjects.Add(15, 2, 200, 200, dummyImage);
                    OleObject ole2 = sheet.OleObjects[oleIndex2];
                    ole2.DisplayAsIcon = false;
                    ole2.ObjectData = dummyExcel;
                    ole2.ProgID = "Excel.Sheet.12";

                    Console.WriteLine("Added two sample OLE objects to the worksheet.");
                }

                // -----------------------------------------------------------------
                // 3. Hide background OLE objects (those displayed as icons).
                // -----------------------------------------------------------------
                foreach (OleObject ole in sheet.OleObjects)
                {
                    try
                    {
                        if (ole.DisplayAsIcon)
                        {
                            // Hide the object using the IsHidden property.
                            ole.IsHidden = true;
                            Console.WriteLine($"OleObject \"{ole.Name}\" (icon) hidden via IsHidden = true.");
                        }
                        else
                        {
                            // Ensure non‑icon objects are visible.
                            ole.IsHidden = false;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing OleObject \"{ole.Name}\": {ex.Message}");
                    }
                }

                // -----------------------------------------------------------------
                // 4. Save the modified workbook.
                // -----------------------------------------------------------------
                string outputPath = "OutputWorkbook.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to \"{outputPath}\" with OLE object visibility updated.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }

        private static byte[] CreateDummyImageBytes()
        {
            // Empty byte array used as placeholder image data.
            return new byte[0];
        }
    }
}
