using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using System.Text.RegularExpressions;

namespace ConsoleApp2
{
    internal class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the PDF document
                Console.Write("Please enter the path : ");
                string inputPath = Console.ReadLine();

                // Define the CSV file path
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string csvFilePath = desktopPath + @"\PDF_files.csv";

                // Check if the path exists
                if (Directory.Exists(inputPath))
                {
                    // Get all .pdf files in the directory
                    string[] dwgFiles = Directory.GetFiles(inputPath, "*.pdf");
                    // Check if there are any .pdf files
                    if (dwgFiles.Length > 0)
                    {
                        // Append results to the CSV file
                        using (StreamWriter writer = new StreamWriter(csvFilePath, true)) // 'true' for append mode
                        {
                            foreach (string file in dwgFiles)
                            {
                                // Text to search for
                                string searchText = "CONTENANCE";

                                // Open the PDF document
                                using (PdfReader reader = new PdfReader(file))
                                using (PdfDocument pdfDoc = new PdfDocument(reader))
                                {
                                    int pageCount = pdfDoc.GetNumberOfPages();
                                    for (int i = 1; i <= pageCount; i++)
                                    {
                                        string pageText = ExtractTextFromPage(pdfDoc, i);
                                        // Check if the search text is in the extracted text
                                        if (pageText.Contains(searchText, StringComparison.OrdinalIgnoreCase))
                                        {
                                            //Start Regex
                                            string pattern = "(?<=(CONTENANCE ADOPTEE  = )).*(?=ca)";
                                            Match match = Regex.Match(pageText, pattern);
                                            if (match.Success)
                                            {
                                                Console.WriteLine(Path.GetFileNameWithoutExtension(file) + " CC : " + match.Value);
                                                // Write the file path and Contenance to the CSV file
                                                writer.WriteLine(Path.GetFileNameWithoutExtension(file) + " CC : " + match.Value);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine("No .pdf files found in the specified directory.");
                    }
                }
                Console.WriteLine("CSV file hase ben created to " + csvFilePath);
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);            
            }
        }

        static string ExtractTextFromPage(PdfDocument pdfDoc, int pageNumber)
        {
            var page = pdfDoc.GetPage(pageNumber);
            var strategy = new SimpleTextExtractionStrategy();
            var text = PdfTextExtractor.GetTextFromPage(page, strategy);
            return text;
        }
    }
}