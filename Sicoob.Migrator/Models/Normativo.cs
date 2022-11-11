using Microsoft.SharePoint.Client;
using Sicoob.Migrator.Properties;
using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;

namespace Sicoob.Migrator.Models
{
    internal class Normativo
    {
        public string Id { get; set; }
        public string Path { get; set; }
        public string FileName { get; set; }

        public Normativo(string id, string path, string fileName)
        {
            Id = id ?? throw new ArgumentNullException(nameof(id));
            Path = path ?? throw new ArgumentNullException(nameof(path));
            FileName = fileName ?? throw new ArgumentNullException(nameof(fileName));
            Path = Path.Replace("~/", string.Empty).Replace('\\', '/');
        }

        public MemoryStream DownloadAttachment(ClientContext ctx, Settings Settings)
        {
            ClientResult<Stream> fileStream;
            try
            {
                Folder folder = Navigate(ctx, Settings, Path.Split('/'));

                ctx.Load(folder.Files);
                ctx.ExecuteQuery();

                var file = folder.Files.First(fl => fl.Name == FileName);

                ctx.Load(file);
                ctx.ExecuteQuery();

                fileStream = file.OpenBinaryStream();

                ctx.ExecuteQuery();
            }
            catch (WebException)
            {
                throw;
            }
            catch (InvalidOperationException)
            {
                try
                {
                    Folder folder = Navigate(ctx, Settings, Path.Split('/'));

                    ctx.Load(folder.Files);
                    ctx.ExecuteQuery();

                    var id = Regex.Replace(Id, @"[A-Za-z]+", "");
                    var file = folder.Files.First(fl => fl.Name == $"{id}_{FileName}");

                    ctx.Load(file);
                    ctx.ExecuteQuery();

                    fileStream = file.OpenBinaryStream();

                    ctx.ExecuteQuery();
                }
                catch (Exception)
                {
                    Program.Log(null, $"Não foi possível fazer download do arquivo '{FileName}'.", Program.LogLevel.Information);
                    throw;
                }
            }

            var ms = new MemoryStream();
            fileStream.Value.CopyTo(ms);

            return ms;
        }

        private Folder Navigate(ClientContext ctx, Settings Settings, string[] paths)
        {
            ctx.Load(ctx.Web.Folders);
            ctx.ExecuteQuery();

            Folder folder = ctx.Web.Folders.GetByPath(ResourcePath.FromDecodedUrl(Settings.Libary));
            ctx.Load(folder);
            ctx.ExecuteQuery();

            foreach (var path in paths)
            {
                folder = folder.Folders.GetByPath(ResourcePath.FromDecodedUrl(path));
                ctx.Load(folder);
                ctx.ExecuteQuery();
            }

            return folder;
        }

        public override bool Equals(object obj)
        {
            if (!(obj is Normativo))
                return false;

            Normativo normativo = obj as Normativo;

            if (normativo.Id != Id ||
                normativo.FileName != FileName ||
                (normativo.Path != Path && normativo.Path != Path + '/'))
                return false;

            return true;
        }
    }
}
