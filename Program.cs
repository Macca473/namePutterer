using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using Spire.Xls;


namespace namePutterer
{
    class Program
    {
        static string[] SearchCells = new string[] { "28D", "E39", "30D", "29D", "31E", "30E" };

        static string[] SearchInt = new string[] { "DA", "AC", "WT", "CT"};

        static string DirOffset = "./TestExcels/";

        public class UserInfo
        {
            public string username { get; set; }
            public string password { get; set; }
        }

        static void Main(string[] args)
        {

           DirectoryInfo d = new DirectoryInfo(DirOffset);

            FileInfo[] xlsxFiles = d.GetFiles("*.xls");

            List<FileInfo> allFiles = d.GetFiles("*").ToList();

            for (int inx = 0; inx < xlsxFiles.Length; inx++)
                {
                    Console.WriteLine(DirOffset + xlsxFiles[inx].Name);

                    if (FindThing(xlsxFiles[inx], out string str))
                    {
                        Console.WriteLine("Found " + str);

                        

                        try
                        {
                            using (FileStream stream = File.Open(DirOffset + xlsxFiles[inx].Name, FileMode.Open))
                            {
                                stream.Close();

                                List<FileInfo> SameName = allFiles.Where(x => Path.GetFileNameWithoutExtension(x.FullName) == Path.GetFileNameWithoutExtension(xlsxFiles[inx].FullName)).ToList();

                                foreach (FileInfo thisFile in SameName)
                                    {
                                        Move(thisFile.Name, str);
                                    }

                                
                            }
                        } catch (IOException ex)
                        {
                            Console.WriteLine(ex);
                        }
                    

                    }
                }
        }

        static bool FindThing(FileInfo file, out string str)
        {

            Workbook workbook = new Workbook();

            workbook.LoadFromFile(DirOffset + file.Name);

            Worksheet sheet = workbook.Worksheets[0];

            foreach (string rng in SearchCells)
            {
                string trialstring = sheet.Range[rng].Text;

                //Console.WriteLine("searching for: " + trialstring);

                if (trialstring == null)
                {

                }
                else if (trialstring.Length >= 7 && trialstring.Length <= 10)
                {
                    foreach (string SearchInt in SearchInt)
                    {
                        if (trialstring.StartsWith(SearchInt))
                        {
                            str = trialstring;

                            return true;
                        }
                    }
                }
            }

            str = "";

            return false;
        }

        static void Move(string fileName, string thing)
        {
            File.Move(DirOffset + fileName, DirOffset + thing + " - " + fileName);
        }
    }
}
