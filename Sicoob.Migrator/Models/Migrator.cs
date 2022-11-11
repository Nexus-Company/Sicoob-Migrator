using Bytescout.Spreadsheet;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using Sicoob.Migrator.Properties;
using System;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Threading;
using System.Threading.Tasks;
using File = Microsoft.SharePoint.Client.File;

namespace Sicoob.Migrator.Models
{
    internal class Migrator : IDisposable
    {
        internal static string tempPath = $@"{Path.GetTempPath()}Sicoob Migrator";
        private static System.Drawing.Color sucessColor = System.Drawing.Color.Green;
        private ClientContext ctx;
        private Worksheet worksheet;
        private Spreadsheet document;
        public Settings Settings { get; set; }
        public string OutPut { get; private set; }

        public string resultPath
        {
            get
            {
                string resultPath = $@"{Environment.CurrentDirectory}\Results";

                if (!Directory.Exists(resultPath))
                    Directory.CreateDirectory(resultPath);

                return resultPath;
            }
        }
        public Migrator(Settings settings)
        {
            Settings = settings;
            document = new Spreadsheet();
        }

        public void Connect()
        {
            var authManager = new AuthenticationManager();
            ctx = authManager.GetWebLoginClientContext(Settings.Endpoint);
            OutPut = $@"{resultPath}\{DateTime.Now:dd_MM_yyyy_HH_mm_ss}.xlsx";
            worksheet = document.Workbook.Worksheets.Add("Resultados");
        }

        /// <summary>
        /// Atualiza o arquivo com referências local.
        /// </summary>
        /// <returns>Retorna o diretório do arquivo</returns>
        public string UpdateExcel()
        {
            Site mySite = ctx.Site;
            Web web = ctx.Web;
            Folder folder = web.Folders.GetByPath(ResourcePath.FromDecodedUrl(Settings.Libary));

            ctx.Load(folder);
            ctx.Load(folder.Files);
            ctx.ExecuteQuery();

            File @base = folder.Files.FirstOrDefault(fs => fs.Name == Settings.Base);

            ctx.Load(@base);
            ctx.ExecuteQuery();

            var fileStream = @base.OpenBinaryStream();

            ctx.ExecuteQuery();

            var ms = new MemoryStream();
            fileStream.Value.CopyTo(ms);

            if (!Directory.Exists(tempPath))
                Directory.CreateDirectory(tempPath);

            string path = $"{tempPath}\\{Settings.Base}";

            if (System.IO.File.Exists(path))
                System.IO.File.Delete(path);

            System.IO.File.WriteAllBytes(path, ms.ToArray());
            return path;
        }

        /// <summary>
        /// Carrega a lista de normativos local
        /// </summary>
        /// <param name="path">diretório do executável local</param>
        /// <returns>Retorna lista de normativos</returns>
        public Normativo[] LoadExcel(string path)
        {
            Spreadsheet document = new Spreadsheet();
            document.LoadFromFile(path);
            Workbook book = document.Workbook;
            Worksheet sheet = book.Worksheets[0];
            var normativos = new System.Collections.Generic.List<Normativo>();
            Thread[] threads = new Thread[Settings.Threads];
            int count = (sheet.UsedRangeRowMax / Settings.Threads);

            for (int i = 0; i < Settings.Threads; i++)
            {
                Thread th = new Thread((object obj) =>
                {
                    Reference @ref = obj as Reference;
                    @ref.Normativos.AddRange(ReadRows(@ref.Worksheet, @ref.Position * @ref.Count, @ref.Count));
                });

                threads[i] = th;
                threads[i].Start(new Reference(count, i, sheet, ref normativos));
            }

            AwaitThreads(threads);
            return normativos.ToArray();
        }

        /// <summary>
        /// Adiciona os anexos a todos os normativos disponiveis na lista.
        /// </summary>
        /// <param name="normativos">Lista de normativos para serem atualizados.</param>
        public async Task UpdateNormativosAsync(Normativo[] normativos)
        {
            Site site = ctx.Site;
            Web web = ctx.Web;
            List list = web.Lists.GetByTitle(Settings.OutPut.Title);

            ctx.Load(list);
            ctx.ExecuteQuery();

            foreach (var item in normativos)
            {
                await Task.Run(() =>
                {
                    try
                    {
                        var query = CamlQuery.CreateAllItemsQuery();
                        query.ViewXml = $"<View><RowLimit>1</RowLimit><Query><Where><Eq><FieldRef Name='idUnicoNormativo' /><Value Type='Text'>{item.Id}</Value></Eq></Where></Query></View>";

                        var items = list.GetItems(query);
                        ctx.Load(items);
                        ctx.ExecuteQuery();

                        ListItem listItem = items.First();

                        ctx.Load(listItem);
                        ctx.ExecuteQuery();

                        ctx.Load(listItem.AttachmentFiles);
                        ctx.ExecuteQuery();

                        if (listItem.AttachmentFiles
                            .FirstOrDefault(fs => fs.FileName == item.FileName) != null)
                        {
                            Attachment oAttachmnt = listItem.AttachmentFiles.GetByFileName(item.FileName);
                            oAttachmnt.DeleteObject();
                            ctx.ExecuteQuery();
                        }

                        string fileName = $@"{tempPath}/{item.FileName}";
                        var ms = item.DownloadAttachment(ctx, Settings);
                        System.IO.File.WriteAllBytes(fileName, ms.ToArray());

                        using (var fs = new FileStream(fileName, FileMode.Open))
                        {
                            var attInfo = new AttachmentCreationInformation();
                            attInfo.FileName = item.FileName;
                            attInfo.ContentStream = fs;

                            var att = listItem.AttachmentFiles.Add(attInfo);
                            ctx.ExecuteQuery();
                        }

                        Reporte(new ResultadoNormativo(item), ref normativos);
                        Program.Log(null, $"Sucesso ao adicionar anexo na normativa '{item.Id}'.", Program.LogLevel.Success);
                    }
                    catch (System.Net.WebException ex)
                    {
                        Reporte(new ResultadoNormativo(item, "Houve um problema na conexão com o servidor."), ref normativos);
                        throw;
                    }
                    catch (Exception ex)
                    {
                        string message = $"Houve um erro ao atualizar a normativa com idUnico: '{item.Id}'";
                        Reporte(new ResultadoNormativo(item, message), ref normativos);
                        Program.Log(ex, message);
                    }
                });
            }
        }

        /// <summary>
        /// Remove os normativos que já foram anexados.
        /// </summary>
        /// <param name="list">Lista de normativos carregados</param>
        /// <returns>Lista com normativos removidos.</returns>
        public Normativo[] RemoveByResults(ref Normativo[] list)
        {
            System.Collections.Generic.List<Normativo> normativos = new System.Collections.Generic.List<Normativo>(list);
            var files = Directory.GetFiles(resultPath);
            Thread[] threads = new Thread[files.Length];

            for (int i = 0; i < files.Length; i++)
            {
                string file = files[i];
                FileInfo info = new FileInfo(file);

                if (info.Extension != ".xlsx")
                    continue;

                threads[i] = new Thread((object obj) =>
                {
                    Reference @ref = obj as Reference;

                    try
                    {
                        var removels = GetSuccessfuls(@ref.FileName);

                        @ref.Normativos.RemoveAll(item => removels.FirstOrDefault(con => item.Equals(con)) != null);
                    }
                    catch (Exception) { }
                });

                threads[i].Start(new Reference(file, ref normativos));
            }

            AwaitThreads(threads);

            return normativos.ToArray();
        }

        /// <summary>
        /// Limpa o objeto atual
        /// </summary>
        public void Dispose()
        {
            Directory.Delete(tempPath, true);
            document.SaveAs(OutPut);
            document.Dispose();
            ctx.Dispose();
            GC.Collect();
        }

        int count = 0;
        private void Reporte(ResultadoNormativo resultado, ref Normativo[] normativos)
        {
            worksheet.Rows.Insert(count, 1);

            var fillColor = resultado.Success ? sucessColor : System.Drawing.Color.Red;
            var row = worksheet.Rows[count];

            row.FontColor = fillColor;

            row[0].Value = resultado.Normativo.Id;
            row[1].Value = resultado.Date;
            row[2].Value = $@"{resultado.Normativo.Path}/{resultado.Normativo.FileName}";
            row[3].Value = resultado.Error;


            count++;
        }

        private Normativo[] ReadRows(Worksheet sheet, int start, int count)
        {
            int blanks = 0;
            var normativos = new System.Collections.Generic.List<Normativo>();

            for (int i = start; i - start < count && i < sheet.UsedRangeRowMax; i++)
            {
                var row = sheet.Rows[i + 1];

                if (string.IsNullOrEmpty(row[0].Value as string) ||
                    string.IsNullOrEmpty(row[1].Value as string) ||
                    string.IsNullOrEmpty(row[2].Value as string))
                {
                    blanks++;

                    if (blanks >= Settings.MaxBlanks)
                        break;

                    continue;
                }
                else
                {
                    blanks = 0;
                }

                try
                {
                    normativos.Add(new Normativo(
                        row[0].Value as string,
                        row[1].Value as string,
                        row[2].Value as string));
                }
                catch (Exception ex)
                {
                    Program.Log(ex, $"Row line '{i}' incorrect!");
                }
            }

            return normativos.ToArray();
        }
        private void AwaitThreads(Thread[] threads)
        {
            bool finished = false;

            while (!finished)
            {
                finished = true;

                for (int i = 0; i < threads.Length; i++)
                {
                    if (threads[i].ThreadState == ThreadState.Running)
                        finished = false;
                }
            }
        }

        private Normativo[] GetSuccessfuls(string file)
        {
            System.Collections.Generic.List<Normativo> normativos = new System.Collections.Generic.List<Normativo>();
            Spreadsheet document = new Spreadsheet();
            document.LoadFromFile(file);
            var sheet = document.Worksheets[0];

            for (int i = 0; i < sheet.UsedRangeRowMax - 1; i++)
            {
                Row row = sheet.Rows[i];

                string path = row[2].Value as string;
                string fileName = path.Replace('\\', '/').Split('/').Last();

                if (row[0].FontColor.ToArgb() == sucessColor.ToArgb())
                    normativos.Add(new Normativo(
                     row[0].Value as string,
                     path.Replace(fileName, string.Empty),
                     fileName
                     ));
            }

            return normativos.ToArray();
        }

        private class Reference
        {
            public int Count { get; set; }
            public int Position { get; set; }
            public Worksheet Worksheet { get; set; }
            public System.Collections.Generic.List<Normativo> Normativos { get; set; }
            public ClientContext Context { get; set; }
            public Normativo[] NormativosArray { get; set; }
            public string FileName { get; set; }
            public Reference(int count, int position, Worksheet sheet, ref System.Collections.Generic.List<Normativo> normativos)
            {
                Count = count;
                Position = position;
                Normativos = normativos;
                Worksheet = sheet;
            }

            public Reference(string file, ref System.Collections.Generic.List<Normativo> normativos)
            {
                FileName = file;
                Normativos = normativos;
            }

            public Reference(ClientContext ctx, Normativo[] normativos)
            {
                Context = ctx;
                NormativosArray = normativos;
            }
        }
    }

    public static class Extensions
    {
        public static void Split<T>(this T[] array, int index, int count, out T[] first)
        {
            first = array.Skip(index)
                .Take(count)
                .ToArray();
        }
    }
}
