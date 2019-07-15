/* Simple application to recursively replace all hyperlinks in every docx file found within a folder structure
 * the program reads the targets from targets.txt
 * 
 * targets.txt must be in the format:
 * https://www.bing.com.au,https://www.google.com.au
 * 
 * The above will replace all instances of bing with google.
 * 
 * Depends on Aspose https://www.aspose.com/
 * A valid aspose liscense is required and not included
 */
using Aspose.Words;
using Aspose.Words.Fields;
using FindAndReplaceHyperlink.model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace FindAndReplaceHyperlink
{
    class Program
    {
        static void Main(string[] args)
        {
            //Set Aspose Lisence 
            SetLicense();
            args = new string[1];
            args[0] = "-replace";

            if (args.Count() == 0 ||
               (args[0].ToLower() != "-report") ||
               (args[0].ToLower() != "-replace"))
            {
                //-report will only output potential matches to a excel file, -replace will update file
                Console.WriteLine("must include -report or -replace");
            }
            else
            {
                var reportOnly = args[0].ToLower() == "-report";
                var rootFolder = AppDomain.CurrentDomain.BaseDirectory;

                //get replace target from targets.txt and populate into a List<string>
                var targetList = ReadTarget();
                var stringList = new List<string>();
                var output = SearchAllFiles(ReplaceModel.Convert(targetList), reportOnly, stringList, rootFolder);
                File.WriteAllLines("output.csv", stringList.ToArray());
            }
        }

        //Set License for Aspose Word.
        //More information here https://docs.aspose.com/display/wordsnet/Licensing
        private static void SetLicense()
        {
            new License().SetLicense("license/aspose.words.lic");
        }

        /// <summary>
        /// Reard targets.txt and return as a List<string>
        /// targets.txt should be in the format:
        /// oldhyperlink,newhyperlink  eg:
        /// https://www.bing.com.au,https://www.google.com.au
        /// </string>
        /// </summary>
        /// <returns></returns>
        private static IEnumerable<string> ReadTarget()
        {
            try
            {
                return File.ReadAllLines("targets.txt");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error", ex.ToString());
            }
            throw new FileNotFoundException("target.txt must exist in folder");
        }


        /// <summary>
        /// Recursively search all files and folders within the director
        /// </summary>
        /// <param name="replaceTargetList"></param>
        /// <param name="reportOnly"></param>
        /// <param name="stringList"></param>
        /// <param name="folder"></param>
        /// <returns></returns>
        private static List<string> SearchAllFiles(List<ReplaceModel> replaceTargetList,
            bool reportOnly,
            List<string> stringList,
            string folder)
        {
            Console.WriteLine($"Searching {folder}");
            //Search all folders recursively  
            var folders = Directory.GetDirectories(folder);
            foreach (var newFolder in folders)
            {
                stringList = SearchAllFiles(replaceTargetList, reportOnly, stringList, newFolder);
            }

            string[] files = Directory.GetFiles(folder, "*.doc*");
            foreach (string file in files)
            {
                Console.WriteLine("Searching " + file);
                Document document = null;

                //Load DOCX using aspose
                if (GetDocument(file, out document))
                {
                    //itterate throguh all "fields"
                    //More informatio here https://docs.aspose.com/display/wordsnet/Working+with+Fields
                    foreach (Field field in document.Range.Fields)
                    {
                        //Only work on Hyperlinks
                        if (field.Type == FieldType.FieldHyperlink)
                        {
                            Console.WriteLine("Found");
                            var link = (FieldHyperlink)field;
                            var replaceModel = replaceTargetList
                                .Where<ReplaceModel>((Func<ReplaceModel, bool>)(a => link.Address.ToLower()
                                .Contains(a.From.ToLower())))
                                .FirstOrDefault<ReplaceModel>();
                            if (replaceModel != null)
                            {
                                var newAddress = Regex.Replace(link.Address, replaceModel.From, replaceModel.To, RegexOptions.IgnoreCase);
                                var outputLine = file + "," + link.Address + "," + newAddress;
                                stringList.Add(outputLine);
                                if (!reportOnly)
                                    link.Address = newAddress;
                            }
                        }
                    }
                    //Save updated docx with Aspose
                    if (!reportOnly)
                    {
                        document.Save(file);
                    }
                }
            }
            return stringList;
        }

        /// <summary>
        /// Read Docx
        /// </summary>
        /// <param name="file"></param>
        /// <param name="document"></param>
        /// <returns></returns>
        private static bool GetDocument(string file, out Document document)
        {
            document = (Document)null;
            try
            {
                document = new Document(file);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.ToString());
            }
            return document != null;
        }
    }
}
