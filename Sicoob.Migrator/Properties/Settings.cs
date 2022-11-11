using Newtonsoft.Json;
using System;
using System.IO;

namespace Sicoob.Migrator.Properties
{
    internal class Settings
    {
        public string Endpoint { get; set; }
        public string Libary { get; set; }
        public string Base { get; set; }
        public int MaxBlanks { get; set; }
        public int SleepTime { get; set; }
        public int Threads { get; set; }
        public OutPut OutPut { get; set; }
        public static Settings Load()
            => JsonConvert.DeserializeObject<Settings>(
                File.ReadAllText(Environment.CurrentDirectory + @"\Properties\settings.json")) ?? new Settings();
    }
    internal class OutPut
    {
        public string Title { get; set; }
    }
}