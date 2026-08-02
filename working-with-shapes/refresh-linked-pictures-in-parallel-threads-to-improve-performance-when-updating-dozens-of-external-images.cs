// Title: C# – Refresh Linked Pictures in Excel with Aspose.Cells Using Parallel Threads
// Description: Sample code that creates an Excel workbook, inserts linked pictures, saves the file, then reloads it and refreshes every linked image in parallel (max 5 concurrent downloads). Each picture is downloaded, replaced with an embedded image, and its original position and size are preserved before the workbook is saved again.
// Keywords: Aspose.Cells C# linked picture refresh | parallel image download Aspose.Cells | replace linked picture with embedded Excel | SemaphoreSlim concurrent downloads | thread‑safe shape update Aspose.Cells | Excel image refresh .NET | US developers Aspose.Cells example | GitHub Aspose.Cells linked picture sample
// Common Searches: how to refresh linked pictures in Aspose.Cells | parallel download of external images for Excel using C# | replace linked picture with embedded picture Aspose.Cells | limit concurrent image downloads Aspose.Cells | thread‑safe picture update in Excel workbook
// Developer Intent: Update all linked pictures in an Excel workbook by downloading their source files concurrently and embedding them while keeping the original layout.
// Use Cases: Speed up report generation that contains dozens of external images. | Convert linked pictures to embedded ones before sharing the workbook to remove external dependencies. | Throttle network traffic when refreshing many images in large Excel files.
// AI Prompts: Write C# code that uses Aspose.Cells to replace linked pictures with embedded images, downloading the sources in parallel and limiting concurrency with SemaphoreSlim. | Show how to keep a picture's original Top, Left, Height, and Width when swapping a linked picture for an embedded picture in Aspose.Cells. | Explain the thread‑safety considerations when modifying Aspose.Cells shapes inside a Parallel.ForEach loop.

using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsLinkedPictureRefresh
{
    // Sample code that creates an Excel workbook, inserts linked pictures, saves the file, then reloads it and refreshes every linked image in parallel (max 5 concurrent downloads). Each picture is downloaded, replaced with an embedded image, and its original position and size are preserved before the workbook is saved again.
    class Program
    {
        // Sample list of image URLs to be linked in the worksheet
        private static readonly List<string> ImageUrls = new List<string>
        {
            "https://picsum.photos/seed/1/200/150",
            "https://picsum.photos/seed/2/200/150",
            "https://picsum.photos/seed/3/200/150",
            "https://picsum.photos/seed/4/200/150",
            "https://picsum.photos/seed/5/200/150"
        };

        static void Main()
        {
            try
            {
                // -----------------------------------------------------------
                // STEP 1: CREATE workbook and add linked pictures
                // -----------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                int startRow = 1;
                int startColumn = 1;
                const int pictureHeight = 150; // pixels
                const int pictureWidth = 200;  // pixels

                foreach (string url in ImageUrls)
                {
                    // Add linked picture and store the source URL in AlternativeText
                    Picture linkedPic = sheet.Shapes.AddLinkedPicture(startRow, startColumn,
                                                                      pictureHeight, pictureWidth, url);
                    linkedPic.AlternativeText = url;

                    // Move to next column for the next picture
                    startColumn += 5;
                }

                string initialFile = "LinkedPictures.xlsx";

                // Ensure the directory exists before saving
                string initialDir = Path.GetDirectoryName(Path.GetFullPath(initialFile));
                if (!Directory.Exists(initialDir))
                {
                    Directory.CreateDirectory(initialDir);
                }

                workbook.Save(initialFile);
                Console.WriteLine($"Workbook with linked pictures saved to '{initialFile}'.");

                // -----------------------------------------------------------
                // STEP 2: LOAD workbook and refresh linked pictures in parallel
                // -----------------------------------------------------------
                if (!File.Exists(initialFile))
                {
                    Console.WriteLine($"Error: File '{initialFile}' not found.");
                    return;
                }

                Workbook loadedWb = new Workbook(initialFile);
                Worksheet loadedSheet = loadedWb.Worksheets[0];
                PictureCollection pictures = loadedSheet.Pictures;

                // Gather linked pictures and their URLs
                var linkedPictures = new List<(Picture pic, string url)>();
                foreach (Picture pic in pictures)
                {
                    if (pic.IsLink && !string.IsNullOrEmpty(pic.AlternativeText))
                    {
                        linkedPictures.Add((pic, pic.AlternativeText));
                    }
                }

                using (HttpClient httpClient = new HttpClient())
                using (SemaphoreSlim semaphore = new SemaphoreSlim(5)) // limit to 5 concurrent downloads
                {
                    Parallel.ForEach(linkedPictures, linkedItem =>
                    {
                        semaphore.Wait();
                        try
                        {
                            // Download image data
                            byte[] imageData = httpClient.GetByteArrayAsync(linkedItem.url).Result;

                            // Modify Aspose.Cells objects inside a lock (not thread‑safe)
                            lock (loadedSheet)
                            {
                                // Preserve original layout
                                int top = linkedItem.pic.Top;
                                int left = linkedItem.pic.Left;
                                int height = linkedItem.pic.Height;
                                int width = linkedItem.pic.Width;

                                // Remove the old linked picture
                                int index = pictures.IndexOf(linkedItem.pic);
                                if (index >= 0)
                                {
                                    pictures.RemoveAt(index);
                                }

                                // Add new embedded picture from the downloaded bytes
                                using (MemoryStream ms = new MemoryStream(imageData))
                                {
                                    // Add at cell (0,0) with zero offsets; we'll set exact position afterwards
                                    Picture newPic = loadedSheet.Shapes.AddPicture(0, 0, 0, 0, ms);
                                    newPic.Top = top;
                                    newPic.Left = left;
                                    newPic.Height = height;
                                    newPic.Width = width;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Failed to refresh picture from '{linkedItem.url}': {ex.Message}");
                        }
                        finally
                        {
                            semaphore.Release();
                        }
                    });
                }

                // Save the refreshed workbook
                string refreshedFile = "LinkedPictures_Refreshed.xlsx";
                loadedWb.Save(refreshedFile);
                Console.WriteLine($"Refreshed workbook saved to '{refreshedFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
