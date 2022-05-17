using Microsoft.VisualBasic.FileIO;
using Mono.Options;
using System;
using System.Collections.Generic;
using System.IO;

namespace MailMerge
{
    class Program
    {
        static void Main(string[] args)
        {
            var shouldShowHelp = false;
            string emailTemplate = "";
            string emailServer = "";
            string emailUsername = "";
            string emailPassword = "";
            string csvFile = "";
            
            var options = new OptionSet {
                { "t|template=", "Email template file.", (string t) => emailTemplate = t },
                { "server=", "SMTP server", (string s) => emailServer = s },
                { "username=", "SMTP username", (string s) => emailUsername = s },
                { "password=", "SMTP password", (string s) => emailPassword = s },
                { "h|help", "show this message and exit", h => shouldShowHelp = h != null },
            };
            List<string> extra;
            try
            {
                extra = options.Parse(args);
                if (extra.Count > 0) csvFile = extra[0];
            }
            catch (OptionException e)
            {
                ShowHelp(options, e.Message);
            }
            if (shouldShowHelp)
            {
                ShowHelp(options);
                return;
            }
            if (csvFile == "") { ShowHelp(options, "Missing csvFile"); return; }
            if (emailTemplate == "") { ShowHelp(options, "Missing template"); return; }
            if (emailServer == "") { ShowHelp(options, "Missing server"); return; }
            if (emailUsername == "") { ShowHelp(options, "Missing username"); return; }
            if (emailPassword == "") { ShowHelp(options, "Missing password"); return; }
            using (TextFieldParser parser = new TextFieldParser(csvFile))
            {
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(",");
                while (!parser.EndOfData)
                {
                    MergeJob job = new MergeJob(parser.ReadFields());
                    try
                    {
                        job.execute(emailTemplate, emailServer, emailUsername, emailPassword);
                    } catch (System.Exception e)
                    {
                        Console.WriteLine(job.ToString() + ",\"" + e.Message + "\"");
                    }

                }
            }
        }
        private static void ShowHelp(OptionSet p, String error = "")
        {
            if (error != "")
            {
                Console.WriteLine("Error: " + error);
                Console.WriteLine();
            }
            Console.WriteLine("Usage: mailmerge.exe [OPTIONS]+ csvFile");
            Console.WriteLine("Send one email per row of csvFile. The CSV file should");
            Console.WriteLine("contain one column with a filename to be attached, then");
            Console.WriteLine("one column with the primary recipient email (To field)");
            Console.WriteLine("followed by a variable number of columns containing");
            Console.WriteLine("secondary recipients (CC). The last column is ignored");
            Console.WriteLine();
            Console.WriteLine("The program outputs a CSV file with the same format, with");
            Console.WriteLine("one row per failed email merge. The last column contains the");
            Console.WriteLine("error message. This file may be fed again into mailmerge.exe");
            Console.WriteLine("Options:");
            p.WriteOptionDescriptions(Console.Out);
        }
    }
}
