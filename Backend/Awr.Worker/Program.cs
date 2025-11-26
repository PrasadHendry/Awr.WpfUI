using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using Awr.Worker.Configuration;
using Awr.Worker.DTOs;
using Awr.Worker.Processors;
using Word = Microsoft.Office.Interop.Word;

namespace Awr.Worker
{
    public class Program
    {
        public static readonly List<Word.Application> ActiveWordApps = new List<Word.Application>();

        private static void OnProcessExit(object sender, EventArgs e)
        {
            foreach (var wordApp in ActiveWordApps.ToArray())
            {
                try { wordApp.Quit(false); }
                catch { }
                finally { Marshal.ReleaseComObject(wordApp); }
            }
        }

        static int Main(string[] args)
        {
            AppDomain.CurrentDomain.ProcessExit += OnProcessExit;
            Console.WriteLine("AWR Worker Started...");

            if (args.Length < 2) return WorkerConstants.FailureExitCode;

            // Args[0] is unused now (was filename), we rely on JSON
            string base64JsonInput = args[1];

            try
            {
                var serializer = new JavaScriptSerializer();
                string json = Encoding.UTF8.GetString(Convert.FromBase64String(base64JsonInput));
                var record = serializer.Deserialize<AwrStampingDto>(json);

                Console.WriteLine($" > Mode: {record.Mode}");
                Console.WriteLine($" > Request: {record.RequestNo}");

                for (int i = 1; i <= WorkerConstants.MaxRetries; i++)
                {
                    try
                    {
                        new DocumentProcessor(record).ProcessRequest();
                        Console.WriteLine(" > Success.");
                        return WorkerConstants.SuccessExitCode;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error (Attempt {i}): {ex.Message}");
                        Thread.Sleep(2000);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Fatal: {ex.Message}");
            }

            return WorkerConstants.FailureExitCode;
        }
    }
}