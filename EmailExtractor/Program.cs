using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Xfinium.Pdf;
using Xfinium.Pdf.Content;

namespace EmailExtractor
{
    class Program
    {
        #region Properties.
        /// <summary>
        /// Message instructing the user to exit.
        /// </summary>
        private static string ExitMessage = "Press [enter] to exit.";
        #endregion

        #region Main Program.
        /// <summary>
        /// Run the main program.
        /// </summary>
        /// <param name="args">List of arguments to run the program with.</param>
        static void Main(string[] args)
        {
            FileStream stream = null;
            var isValidPath = false;

            do
            {
                Console.WriteLine("What is the absolute path to the PDF file?");

                var path = Console.ReadLine();

                Console.WriteLine("Reading the file...");

                try
                {
                    stream = File.OpenRead(path);
                }
                catch (FileNotFoundException)
                {
                    Console.WriteLine("Invalid path, file not found in that location.");
                }
                catch (DirectoryNotFoundException)
                {
                    Console.WriteLine("Invalid path, file not found in that location.");
                }
                catch (Exception e)
                {
                    ExitProgram($"Error reading the file: {e.Message}");
                    isValidPath = true;
                }
            }
            while (!isValidPath && stream == null);

            Console.WriteLine("Opening the document...");

            PdfFixedDocument document = null;

            try
            {
                document = new PdfFixedDocument(stream);
            }
            catch (Exception e)
            {
                ExitProgram($"Error opening the document: {e.Message}");
            }

            var context = new PdfContentExtractionContext();
            var documentText = "";

            Console.WriteLine("Parsing the page content...");

            var pageCount = 1;

            foreach (var page in document.Pages)
            {
                try
                {
                    documentText += new PdfContentExtractor(page).ExtractText(context);
                }
                catch (Exception e)
                {
                    ExitProgram($"Unable to extract content from page #{pageCount}: {e.Message}");
                }

                pageCount++;
            }

            Console.WriteLine("Generating the output file...");

            var isValidOutputType = false;

            Console.WriteLine("Parsing Completed. How would you like the output? [csv, txt]");

            do
            {
                var outputType = Console.ReadLine();

                switch (outputType)
                {
                    case "csv":
                        isValidOutputType = true;
                        GenerateCsv(GetSavePath(), GetSaveName(), documentText);
                        break;
                    case "txt":
                        isValidOutputType = true;
                        GenerateTxt(GetSavePath(), GetSaveName(), documentText);
                        break;
                    default:
                        Console.WriteLine("Unrecognized output type. Please select one of the following: csv, txt.");
                        break;
                }
            }
            while (!isValidOutputType);

            ExitProgram("Parsing Complete.");
        }
        #endregion

        #region Exit Methods.
        /// <summary>
        /// Writes the exit message for when the program is ending, with
        /// an optional prefix message.
        /// </summary>
        /// <param name="message">Final message to write before telling the user they can exit.<param>
        private static void ExitProgram(string message = null)
        {
            if (!string.IsNullOrEmpty(message))
            {
                Console.WriteLine(message);
            }

            Console.WriteLine(ExitMessage);
            Console.ReadLine();

            Environment.Exit(0);
        }
        #endregion

        #region User Input Methods.
        /// <summary>
        /// Get the path to save from the user.
        /// </summary>
        /// <returns>The save path.</returns>
        private static string GetSavePath()
        {
            Console.WriteLine("Where would you like to save the output?");
            return Console.ReadLine();
        }

        /// <summary>
        /// Get the file name to save from the user.
        /// </summary>
        /// <returns>The save name.</returns>
        private static string GetSaveName()
        {
            Console.WriteLine("What name would you like to save the output with?");
            return Console.ReadLine();
        }
        #endregion

        #region Output Methods.
        /// <summary>
        /// Generate a CSV file.
        /// </summary>
        /// <param name="path">Path to save to.</param>
        /// <param name="name">Name to save with.</param>
        /// <param name="content">Content to convert to CSV format.</param>
        private static void GenerateCsv(string path, string name, string content)
        {
            // Create the two-dimensional array that will represent the CSV file,
            // and being with the list of headers.
            var csv = new List<List<string>>
            {
                new List<string> { "Name", "Company", "Title", "Phone", "Mobile", "Email", "Address" }
            };

            // Split the content by lines.
            var lines = content.Split('\n');

            var isHeader = true;
            var isAddress = false;
            var keywords = new Dictionary<string, int>
            {
                {"Mr.", 0},
                {"Ms.", 0},
                {"Mrs.", 0},
                {"Dr.", 0},
                {"Title:", 2},
                {"Phone:", 3},
                {"Mobile:", 4},
                {"Email:", 5}
            };
            var keywordKeys = keywords.Keys.ToList();

            var namePlacement = csv.First().IndexOf("Name");
            var companyPlacement = csv.First().IndexOf("Company");
            var addressPlacement = csv.First().IndexOf("Address");

            var pageNumberExpression = new Regex("^(\\d{1}|\\d{2}|\\d{3}|\\d{4})$");

            foreach (var line in lines)
            {
                var lineTrimmed = line.Trim();

                // If the trimmed line matched the page number expression,
                // we will skip the line because it just represents the end
                // of a page.
                if (pageNumberExpression.IsMatch(lineTrimmed))
                {
                    continue;
                }

                var match = keywordKeys.FirstOrDefault(keyword => lineTrimmed.StartsWith(keyword));

                if (!string.IsNullOrEmpty(match))
                {
                    if (isHeader || isAddress)
                    {
                        // Make sure the address is surrounded in double qoutes, so
                        // commans that may exist don't interfere with the CSV format.
                        csv.Last()[addressPlacement] = $"\"{csv.Last()[addressPlacement]}\"";

                        csv.Add(new List<string> { "", "", "", "", "", "", "" });

                        isAddress = false;
                        isHeader = false;
                    }

                    var placement = keywords[match];
                    var lineValue = lineTrimmed;

                    if (lineValue.IndexOf(':') > -1)
                    {
                        lineValue = lineValue.Substring(lineValue.IndexOf(':') + 1).Trim();
                    }

                    // Make sure the value is surrounded in double qoutes, so
                    // commans that may exist don't interfere with the CSV format.
                    csv.Last()[placement] = $"\"{lineValue}\"";
                }
                else if (!isHeader)
                {
                    if (!isAddress)
                    {
                        // If we haven't started collecting address information, and we
                        // still have collected the name, we can assume the current content
                        // is the company name data.
                        if (string.IsNullOrEmpty(csv.Last()[namePlacement]))
                        {
                            csv.Last()[companyPlacement] = $"\"{lineTrimmed}\"";
                        }
                        else
                        {
                            isAddress = true;
                        }
                    }
                    else
                    {
                        // If we have already been reading into the address field
                        // we need to separate the content with a space.
                        csv.Last()[addressPlacement] += " ";
                    }

                    csv.Last()[addressPlacement] += lineTrimmed;
                }
            }

            // Make sure the address is surrounded in double qoutes, so
            // commans that may exist don't interfere with the CSV format.
            csv.Last()[addressPlacement] = $"\"{csv.Last()[addressPlacement]}\"";

            WriteOutput(path, name, "csv", 
                        string.Join("\n", csv.Select(row => string.Join(",", row))));
        }

        /// <summary>
        /// Generates a TXT file.
        /// </summary>
        /// <param name="path">Path to save to.</param>
        /// <param name="name">Name to save with.</param>
        /// <param name="content">Content to convert to CSV format.</param>
        private static void GenerateTxt(string path, string name, string content)
        {
            WriteOutput(path, name, "txt", content);
        }

        private static void WriteOutput(string path, string name, string extension, string content)
        {
            try
            {
                File.WriteAllText($"{path}/{name}.{extension}", content);
            }
            catch (Exception e)
            {
                ExitProgram($"Unable to generate output file: {e.Message}");
            }
        }
        #endregion
    }
}
