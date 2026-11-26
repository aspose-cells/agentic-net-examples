// Title: How to replace the source Word document of an embedded OLE object and modify its properties using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# change embedded OLE object's source file to a different Word document | update OLE object label and AutoLoad flag in an Excel workbook with Aspose.Cells | replace OleObject.ObjectData with a new .docx file programmatically in C# | set OleObject.ObjectSourceFullName to a new path using Aspose.Cells .NET | edit properties of an embedded OLE object after loading a workbook in C#
// Tags: replace embedded OLE object source Aspose.Cells | update OleObject.ObjectData docx C# | modify OleObject label AutoLoad property | set OleObject.ObjectSourceFullName path | edit OLE object properties Excel Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace OleObjectEditor
{
    // The example creates an Excel workbook, embeds a Word document as an OLE object with a preview image, saves it, then reloads the workbook, replaces the OLE object's source file path and embedded data with a new Word document, updates its label, display settings, and AutoLoad behavior, and finally saves the modified workbook as a new file.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths – adjust these to point to real files on your machine.
            string excelPath = "OleObjectDemo.xlsx";
            string initialWordPath = "InitialDocument.docx";
            string newWordPath = "UpdatedDocument.docx";
            string previewImagePath = "preview.png";

            // Ensure the preview image exists; otherwise create a minimal 1x1 PNG.
            if (!File.Exists(previewImagePath))
            {
                try
                {
                    // Base64 for a 1x1 transparent PNG.
                    const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2ZcAAAAASUVORK5CYII=";
                    byte[] pngBytes = Convert.FromBase64String(base64Png);
                    File.WriteAllBytes(previewImagePath, pngBytes);
                    Console.WriteLine($"Created dummy preview image: {previewImagePath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to create dummy preview image: {ex.Message}");
                }
            }

            // Ensure the initial and new Word documents exist; otherwise create dummy files.
            CreateDummyFileIfMissing(initialWordPath, "This is the original Word document.");
            CreateDummyFileIfMissing(newWordPath, "This is the updated Word document.");

            try
            {
                // -----------------------------------------------------------------
                // 1. Create a workbook and add an OLE object (embedded) for demo.
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Read the preview image into a byte array.
                byte[] imageBytes = File.ReadAllBytes(previewImagePath);

                // Add the OLE object with the preview image.
                int oleIndex = sheet.OleObjects.Add(2, 2, 120, 120, imageBytes);
                OleObject ole = sheet.OleObjects[oleIndex];

                // Set the embedded object's data to the initial Word document.
                ole.ObjectData = File.ReadAllBytes(initialWordPath);
                ole.ObjectSourceFullName = Path.GetFullPath(initialWordPath);
                ole.DisplayAsIcon = true;
                ole.Label = "Original Document";
                ole.FileFormatType = FileFormatType.Docx;
                ole.IsLink = false; // Embedded, not linked.

                // Save the workbook with the original OLE object.
                workbook.Save(excelPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook created with OLE object and saved to '{excelPath}'.");

                // -----------------------------------------------------------------
                // 2. Load the workbook again and edit the OLE object's source.
                // -----------------------------------------------------------------
                if (!File.Exists(excelPath))
                {
                    Console.WriteLine($"Workbook file not found: {excelPath}");
                    return;
                }

                Workbook loadedWb = new Workbook(excelPath);
                Worksheet loadedSheet = loadedWb.Worksheets[0];

                if (loadedSheet.OleObjects.Count == 0)
                {
                    Console.WriteLine("No OLE objects found in the workbook.");
                    return;
                }

                OleObject oleToEdit = loadedSheet.OleObjects[0];

                // Update the source file path and embedded data.
                oleToEdit.ObjectSourceFullName = Path.GetFullPath(newWordPath);
                oleToEdit.ObjectData = File.ReadAllBytes(newWordPath);
                oleToEdit.Label = "Updated Document";
                oleToEdit.FileFormatType = FileFormatType.Docx;
                oleToEdit.DisplayAsIcon = true;
                oleToEdit.AutoLoad = true;   // Load automatically when workbook opens.
                oleToEdit.AutoUpdate = false; // Do not auto‑update linked data.

                // Save the modified workbook.
                string updatedExcelPath = "OleObjectDemo_Updated.xlsx";
                loadedWb.Save(updatedExcelPath, SaveFormat.Xlsx);
                Console.WriteLine($"OLE object edited. Updated workbook saved to '{updatedExcelPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }

        private static void CreateDummyFileIfMissing(string path, string content)
        {
            if (!File.Exists(path))
            {
                try
                {
                    File.WriteAllText(path, content);
                    Console.WriteLine($"Created dummy file: {path}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to create dummy file '{path}': {ex.Message}");
                }
            }
        }
    }
}
