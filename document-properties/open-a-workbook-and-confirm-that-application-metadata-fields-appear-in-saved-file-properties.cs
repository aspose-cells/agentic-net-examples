// Title: How to set and validate built‑in Excel document properties with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code using Aspose.Cells to assign Application, Author, Title, Subject, and Keywords built‑in document properties to a workbook and save it. | Write C# that loads the saved workbook, reads the same built‑in document properties via the BuiltInDocumentProperties indexer, and stores them in variables. | Create a C# verification routine that compares the retrieved property values with the expected strings and prints a success or failure message.
// Common Searches: Aspose.Cells C# set built‑in document properties like Application and Author | C# verify Excel metadata persistence after saving with Aspose.Cells | Read built‑in document properties from an existing .xlsx using Aspose.Cells .NET | How to programmatically add keywords to an Excel file with Aspose.Cells | Check if custom application name is stored in workbook properties using Aspose.Cells
// Tags: Aspose.Cells set built-in document properties | Aspose.Cells read workbook metadata | C# verify Excel property persistence | BuiltInDocumentProperties indexer usage | Excel file metadata verification Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads or creates an Excel workbook, uses the BuiltInDocumentProperties indexer to set Application, Author, Title, Subject, and Keywords, saves the file, reloads it, reads the same properties back, and prints a verification result confirming that the metadata persisted correctly.
class Program
{
    static void Main()
    {
        // Paths for input and output workbooks
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Load existing workbook or create a new one
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets[0].Cells["A1"].PutValue("Sample Data");
            }

            // Set built‑in document properties using the indexer
            var props = workbook.BuiltInDocumentProperties;
            props["Application"].Value = "MyApp";
            props["Author"].Value = "John Doe";
            props["Title"].Value = "Test Workbook";
            props["Subject"].Value = "Metadata Verification";
            props["Keywords"].Value = "Aspose,Metadata";

            // Save the workbook
            workbook.Save(outputPath);

            // Reload to verify metadata persistence
            Workbook reloaded = new Workbook(outputPath);
            var reloadedProps = reloaded.BuiltInDocumentProperties;

            // Retrieve property values using the indexer
            string app = reloadedProps["Application"]?.Value?.ToString();
            string author = reloadedProps["Author"]?.Value?.ToString();
            string title = reloadedProps["Title"]?.Value?.ToString();
            string subject = reloadedProps["Subject"]?.Value?.ToString();
            string keywords = reloadedProps["Keywords"]?.Value?.ToString();

            // Output metadata values
            Console.WriteLine("Application: " + app);
            Console.WriteLine("Author: " + author);
            Console.WriteLine("Title: " + title);
            Console.WriteLine("Subject: " + subject);
            Console.WriteLine("Keywords: " + keywords);

            // Simple verification
            bool isValid = app == "MyApp" &&
                           author == "John Doe" &&
                           title == "Test Workbook" &&
                           subject == "Metadata Verification" &&
                           keywords == "Aspose,Metadata";

            Console.WriteLine(isValid ? "Metadata verification succeeded." : "Metadata verification failed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
