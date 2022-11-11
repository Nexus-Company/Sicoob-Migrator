using Bytescout.Spreadsheet;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using Sicoob.Migrator.Models;
using Sicoob.Migrator.Properties;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using File = Microsoft.SharePoint.Client.File;

namespace Sicoob.Migrator
{
    internal class Program
    {
        private static string logPath = $@"{Environment.CurrentDirectory}\Logs";
        static Settings Settings;
        public static async Task Main(string[] args)
        {
            Settings = Settings.Load();
            string outPut = string.Empty;
            using (Models.Migrator mi = new Models.Migrator(Settings))
            {
                mi.Connect();
                Log(null, $"Sucesso ao connectar com o servidor!", LogLevel.Success, false);
                string path = mi.UpdateExcel();
                Log(null, $"Sucesso ao obter nova rodada da planilha.", LogLevel.Success, false);
                var list = mi.LoadExcel(path);
                list = mi.RemoveByResults(ref list);

                do
                {
                    try
                    {
                        Log(null, $"Lista de normativas carregada!", LogLevel.Information, false);
                        await mi.UpdateNormativosAsync(list);
                        Log(null, $"Termino da atualização de novas inserções.", LogLevel.Information, false);
                    }
                    catch (System.Net.WebException ex)
                    {
                        Log(null, "E necessário esperar para poder continuar!", LogLevel.Information, false);
                        Thread.Sleep(Settings.SleepTime * 100);
                    }
                    catch (Exception ex)
                    {
                        Log(ex, "O serviço teve um erro inesperado");
                    }

                    outPut = mi.OutPut;

                    Process.Start(outPut);
                    Console.WriteLine("\n\nContinuar...");
                    Console.ReadLine();
                } while (true);
            }
        }

        /// <summary>
        /// Guarda no arquivo de despejo para depurar melhor.
        /// </summary>
        /// <param name="ex">Exceção gerada</param>
        /// <param name="message">Mensagem</param>
        /// <param name="level">Nível da mensagem.</param>
        public static void Log(Exception ex, string message, LogLevel level = LogLevel.Error, bool saveFile = true)
        {
            if (!Directory.Exists(logPath))
                Directory.CreateDirectory(logPath);

            string path = $@"{logPath}\{DateTime.Now:yyyy-MM-dd HH}.log";
            string levelName = $" {Enum.GetName(typeof(LogLevel), level)} ";
            message =
                $"\t\t{DateTime.Now:T}" +
                $"\t\t{message}" +
                $"\t\t{(ex != null ? ex.ToString() : string.Empty)}";

            switch (level)
            {
                case LogLevel.Error:
                    Console.BackgroundColor = ConsoleColor.DarkRed;
                    break;
                case LogLevel.Information:
                    Console.BackgroundColor = ConsoleColor.DarkYellow;
                    break;
                case LogLevel.Success:
                    Console.BackgroundColor = ConsoleColor.Green;
                    break;
            }

            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine(levelName);
            Console.BackgroundColor = ConsoleColor.Black;
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine(message);

            message = levelName + message;

            if (!saveFile)
                return;

            if (!System.IO.File.Exists(path))
            {
                System.IO.File.WriteAllText(path, message);
                return;
            }

            try
            {
                message = $"{System.IO.File.ReadAllText(path)}\n\n{message}";

                System.IO.File.WriteAllText(path, message);
            }
            catch (Exception)
            {
            }
        }

        public enum LogLevel
        {
            Error,
            Information,
            Success
        }
    }
}
