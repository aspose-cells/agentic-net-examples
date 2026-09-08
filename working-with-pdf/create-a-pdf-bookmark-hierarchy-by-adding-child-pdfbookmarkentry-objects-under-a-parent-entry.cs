// Title: Create a nested PDF bookmark hierarchy in Aspose.Cells for .NET using reflection
// AI Prompts: Write C# code that uses reflection to instantiate a parent PdfBookmark, add child PdfBookmark entries, and assign the hierarchy to PdfSaveOptions. | Show how to build a PDF bookmark tree for an Excel workbook and save it as PDF without directly referencing the Aspose.Cells.Pdf assembly. | Demonstrate adding child bookmarks to a parent PdfBookmark collection via reflection and exporting the workbook with the bookmark structure.
// Common Searches: how to build PDF bookmark tree from Excel using Aspose.Cells reflection | add child PdfBookmark objects to parent bookmark in C# Aspose.Cells | save workbook as PDF with nested bookmarks without Aspose.Cells.Pdf assembly | Aspose.Cells create hierarchical PDF bookmarks programmatically | C# reflection example for PdfSaveOptions PdfBookmarks property
// Tags: Aspose.Cells PDF bookmarks via reflection | PdfSaveOptions nested bookmark hierarchy | C# parent child PdfBookmark objects | Excel to PDF with bookmark tree | Aspose.Cells.Pdf bookmark collection manipulation

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;

// The example creates an Excel workbook, uses reflection to instantiate a parent PdfBookmark and two child PdfBookmark objects, adds the children to the parent, assigns the bookmark hierarchy to PdfSaveOptions, and saves the workbook as a PDF with a nested bookmark structure.
class PdfBookmarkExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add some data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Data");
            sheet.Cells["A2"].PutValue("More Data");

            // Prepare PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Try to create PDF bookmarks using reflection (avoids direct dependency on Aspose.Cells.Pdf assembly)
            Type pdfBookmarkType = Type.GetType("Aspose.Cells.Pdf.PdfBookmark, Aspose.Cells.Pdf");
            if (pdfBookmarkType != null)
            {
                // Create parent bookmark (title, page number)
                object parentBookmark = Activator.CreateInstance(pdfBookmarkType, "Parent Bookmark", 1);

                // Create child bookmarks
                object childBookmark1 = Activator.CreateInstance(pdfBookmarkType, "Child Bookmark 1", 1);
                object childBookmark2 = Activator.CreateInstance(pdfBookmarkType, "Child Bookmark 2", 1);

                // Add child bookmarks to the parent
                PropertyInfo childBookmarksProp = pdfBookmarkType.GetProperty("ChildBookmarks");
                object childCollection = childBookmarksProp.GetValue(parentBookmark);
                MethodInfo addMethod = childCollection.GetType().GetMethod("Add");
                addMethod.Invoke(childCollection, new[] { childBookmark1 });
                addMethod.Invoke(childCollection, new[] { childBookmark2 });

                // Assign the bookmark hierarchy to the PDF options
                PropertyInfo pdfBookmarksProp = typeof(PdfSaveOptions).GetProperty("PdfBookmarks");
                Array bookmarksArray = Array.CreateInstance(pdfBookmarkType, 1);
                bookmarksArray.SetValue(parentBookmark, 0);
                pdfBookmarksProp.SetValue(pdfOptions, bookmarksArray);
            }
            else
            {
                Console.WriteLine("Aspose.Cells.Pdf assembly not found. PDF will be saved without bookmarks.");
            }

            // Save the workbook as a PDF with (or without) the defined bookmark hierarchy
            string outputPath = "OutputWithBookmarks.pdf";
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
