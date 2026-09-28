// Title: Sign an Excel VBA project from a MemoryStream and save the signed workbook to another Stream using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook from a Stream, applies password protection to its VBA project, and writes the updated workbook to a different Stream with Aspose.Cells. | Show how to use Aspose.Cells to protect a VBA macro in a workbook loaded from a MemoryStream and output the signed workbook to a MemoryStream in .NET.
// Common Searches: Aspose.Cells C# protect VBA project from a stream | How to sign VBA macro in an Excel file using Aspose.Cells .NET | Save signed Excel workbook to MemoryStream with Aspose.Cells | Load workbook from MemoryStream and apply VBA password protection in C#
// Tags: Aspose.Cells VBA project protection API | load workbook from stream Aspose.Cells | save workbook to memory stream Aspose.Cells | protect VBA macro password .NET | sign Excel VBA project using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads a workbook from an input Stream, protects any existing VBA project with a password, saves the workbook as XLSX to an output Stream, and resets the output Stream position for further use.
    public class VbaSigner
    {
        /// <param name="inputStream">Stream containing the original workbook.</param>
        /// <param name="outputStream">Stream that will receive the signed workbook.</param>
        /// <param name="password">Password used to protect the VBA project.</param>
        public void SignVbaProject(Stream inputStream, Stream outputStream, string password)
        {
            try
            {
                // Load the workbook from the input stream
                var workbook = new Workbook(inputStream);

                // Protect the VBA project with the supplied password, if a VBA project exists
                if (workbook.VbaProject != null)
                {
                    // The Protect method requires a flag indicating that a password is required
                    workbook.VbaProject.Protect(true, password);
                }

                // Save the workbook (with protected VBA) to the output stream in XLSX format
                workbook.Save(outputStream, SaveFormat.Xlsx);

                // Reset the output stream position for further reading
                if (outputStream.CanSeek)
                {
                    outputStream.Position = 0;
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error signing VBA project: {ex.Message}");
                throw;
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "signed_output.xlsx";
            const string vbaPassword = "MySecretPassword";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.Error.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                using (var inputStream = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
                using (var outputStream = new MemoryStream())
                {
                    var signer = new VbaSigner();
                    signer.SignVbaProject(inputStream, outputStream, vbaPassword);

                    // Write the signed workbook to the output file
                    using (var fileOutput = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                    {
                        outputStream.CopyTo(fileOutput);
                    }

                    Console.WriteLine($"Signed workbook saved to: {outputPath}");
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
