using JsonFromOrToExcel;
using JsonFromOrToExcel.Objects;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Data.OleDb;
using System.IO;
using System.Linq;

namespace ExcelToJson
{
    class Program
    {
        //Change filesPath To Your Directory Of Json Files
        //Change destinationPath To Your Directory Of Json Files

        static string filesFormat = "*.xlsx";
        static string filesPath = @"YOUR DRIVE>:\<DIRECTORY IN THIS FORMAT: C:\AA\BB\CC\";
        static string destinationPath = @"YOUR DRIVE>:\<DIRECTORY IN THIS FORMAT: C:\AA\BB\CC\";
        static string sheetName = "Sheet";
        static void Main(string[] args)
        {

            Console.WriteLine("Starting");

            Console.WriteLine("FilesPath: {0}", filesPath);
            Console.WriteLine("Files Format: {0}", filesFormat);

            if (filesPath == @"<YOUR DRIVE>:\<DIRECTORY IN THIS FORMAT: C:\AA\BB\CC\")
            {
                Console.WriteLine("Your filesPath Not Correct... Please Put Your Root Path In --filesPath-- Variable");
                return;
            }

            DirectoryInfo d = new DirectoryInfo(filesPath); //This is your Folder

            FileInfo[] filePaths = d.GetFiles(filesFormat); //Getting Excel files

            foreach (var file in filePaths)
            {
                sheetName = Path.GetFileNameWithoutExtension(file.Name);
                Console.WriteLine("Reading File: {0}", sheetName);

                ExcelToJson(file);
            }

        }
        public static void ExcelToJson(FileInfo pathToExcel)
        {
            try
            {
                string fileName = pathToExcel.Name;
                //This connection string works if you have Office 2007+ installed and your 
                //data is saved in a .xlsx file
                var connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0 Xml;HDR=YES""", pathToExcel.FullName);

                //Creating and opening a data connection to the Excel sheet 
                using (var conn = new OleDbConnection(connectionString))
                {
                    conn.Open();

                    var cmd = conn.CreateCommand();
                    cmd.CommandText = String.Format(@"SELECT * FROM [{0}$]", sheetName);

                    using (var rdr = cmd.ExecuteReader())
                    {
                        JsonObject jsonObject = new JsonObject();
                        int rowNumber = 0;
                        //LINQ query - when executed will create anonymous objects for each row
                        var query =
                            (from DbDataRecord row in rdr
                             select row).Select(x =>
                             {
                                 rowNumber++;
                                 Vocab vocab = new Vocab();
                                 vocab.ID = rowNumber;
                                 vocab.English = x[0].ToString().Trim();
                                 vocab.Persian = x[1].ToString().Trim();
                                 return vocab;

                             }).ToList();

                        //Generates JSON from the LINQ query
                        var json = JsonConvert.SerializeObject(query);

                        if (destinationPath == @"<YOUR DRIVE>:\<DIRECTORY IN THIS FORMAT: C:\AA\BB\CC\")
                        {
                            Console.WriteLine("Your destinationPath Not Correct... Please Put Your Destination Path In --destinationPath-- Variable");
                            return;
                        }

                        File.WriteAllText(destinationPath + sheetName + ".json", json);


                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("ExportToJson: Json file could not be saved! Check filepath.\n"
                + ex.Message);
            }

        }
        public static List<Vocab> ToList(List<Vocab> vocabsList)
        {
            var vocabs = new List<Vocab>();

            foreach (var vocab in vocabsList)
            {
                vocabs.Add(new Vocab { English = vocab.English.Trim(), Persian = vocab.Persian.Trim() });
            }
            return vocabs;
        }
    }

}
