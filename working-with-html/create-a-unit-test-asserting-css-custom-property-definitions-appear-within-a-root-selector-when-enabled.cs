// Title: C# Unit Test: Verify :root selector with EnableCssCustomProperties in Aspose.Cells HTML export
// Description: Creates a temporary PNG, adds it twice to a workbook, saves the workbook to HTML with HtmlSaveOptions.EnableCssCustomProperties set to true and false, and asserts that a :root selector appears only when the option is enabled. Cleans up the temporary image after execution.
// Keywords: Aspose.Cells | C# | HtmlSaveOptions | EnableCssCustomProperties | unit test | HTML export | CSS custom properties | root selector | image embedding | memory stream
// Common Searches: Aspose.Cells unit test for CSS custom properties | EnableCssCustomProperties :root selector example | C# test HTML export Aspose.Cells | Verify CSS variables in Aspose.Cells HTML output | How to assert :root selector in saved HTML
// Developer Intent: Confirm that enabling CSS custom properties adds a :root block to the generated HTML and that disabling it removes the block.
// Use Cases: Automated regression testing of Aspose.Cells HTML export for design‑system compliance. | CI pipeline check to ensure CSS variables are emitted only when required. | Reference implementation for developers writing their own unit tests around HtmlSaveOptions.
// AI Prompts: Generate an MSTest method that asserts the presence of a :root selector when EnableCssCustomProperties is true. | Create a reusable helper that saves a Workbook to an HTML string with configurable HtmlSaveOptions for testing. | Write an xUnit test that verifies the :root selector is absent when EnableCssCustomProperties is false.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // Creates a temporary PNG, adds it twice to a workbook, saves the workbook to HTML with HtmlSaveOptions.EnableCssCustomProperties set to true and false, and asserts that a :root selector appears only when the option is enabled. Cleans up the temporary image after execution.
    class Program
    {
        static void Main()
        {
            string tempImagePath = Path.Combine(Path.GetTempPath(), $"tempImg_{Guid.NewGuid()}.png");
            try
            {
                // Create a minimal 1x1 PNG image (transparent)
                byte[] pngBytes = new byte[]
                {
                    0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
                    0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
                    0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,
                    0x08,0x06,0x00,0x00,0x00,0x1F,0x15,0xC4,
                    0x89,0x00,0x00,0x00,0x0A,0x49,0x44,0x41,
                    0x54,0x78,0x9C,0x63,0x60,0x00,0x00,0x00,
                    0x02,0x00,0x01,0xE2,0x21,0xBC,0x33,0x00,
                    0x00,0x00,0x00,0x49,0x45,0x4E,0x44,0xAE,
                    0x42,0x60,0x82
                };

                // Write the PNG to a temporary file
                File.WriteAllBytes(tempImagePath, pngBytes);

                // Ensure the image file exists
                if (!File.Exists(tempImagePath))
                    throw new FileNotFoundException("Temporary image was not created.", tempImagePath);

                // Create a new workbook and add sample data
                var workbook = new Workbook();
                var worksheet = workbook.Worksheets[0];
                worksheet.Cells["A1"].PutValue("Sample Text");

                // Add the same image twice to trigger CSS custom property generation
                worksheet.Pictures.Add(1, 1, tempImagePath);
                worksheet.Pictures.Add(5, 5, tempImagePath);

                // Save with EnableCssCustomProperties = true
                var optionsEnabled = new HtmlSaveOptions
                {
                    EnableCssCustomProperties = true
                };
                string htmlWithCustomProps = SaveWorkbookToHtml(workbook, optionsEnabled);

                // Verify that a :root selector is present
                if (!htmlWithCustomProps.Contains(":root"))
                    throw new Exception("HTML should contain a :root selector when CSS custom properties are enabled.");

                // Save with EnableCssCustomProperties = false
                var optionsDisabled = new HtmlSaveOptions
                {
                    EnableCssCustomProperties = false
                };
                string htmlWithoutCustomProps = SaveWorkbookToHtml(workbook, optionsDisabled);

                // Verify that the :root selector is not present
                if (htmlWithoutCustomProps.Contains(":root"))
                    throw new Exception("HTML should not contain a :root selector when CSS custom properties are disabled.");

                Console.WriteLine("Test passed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
            finally
            {
                // Clean up the temporary image file
                if (File.Exists(tempImagePath))
                {
                    try
                    {
                        File.Delete(tempImagePath);
                    }
                    catch
                    {
                        // Ignore any errors during cleanup
                    }
                }
            }
        }

        private static string SaveWorkbookToHtml(Workbook workbook, HtmlSaveOptions options)
        {
            try
            {
                using (var ms = new MemoryStream())
                {
                    workbook.Save(ms, options);
                    ms.Position = 0;
                    using (var reader = new StreamReader(ms))
                    {
                        return reader.ReadToEnd();
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Failed to save workbook to HTML.", ex);
            }
        }
    }
}
