using Microsoft.Office.Interop.Outlook;
using System;
using System.IO;
using System.Linq;

namespace OutlookRobot
{
    class Program
    {
        static readonly string basePath = @"c:\temp\attemails\";
        static void Main(string[] args)
        {
            try
            {
                Application Application = new Application();
                Folder folder = Application.Session.DefaultStore.GetRootFolder() as Folder;
                GetFolders(folder);
            }
            catch (System.Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine(e);
            }

            Console.ReadLine();
        }

        static void GetFolders(Folder folder)
        {
            Folders childFolders = folder.Folders;
            if (childFolders.Count > 0)
            {
                foreach (Folder childFolder in childFolders)
                {
                    if (childFolder.FolderPath.Contains("RELSCO"))
                    {
                        Console.WriteLine("Estou em: " + folder.FolderPath);
                        GetFolders(childFolder);
                    }
                }
            }

            Console.WriteLine("Estou em: " + folder.FolderPath);
            GetItems(folder);
        }

        static void GetItems(Folder folder)
        {
            string[] ext = { ".xlx", ".xlsx" };

            var fi = folder.Items;

            if (fi != null)
            {
                try
                {
                    foreach (object item in fi)
                    {
                        MailItem mi = (MailItem)item;
                        var att = mi.Attachments;

                        if (att.Count != 0)
                        {
                            if (!Directory.Exists(basePath + folder.FolderPath))
                            {
                                Directory.CreateDirectory(basePath + folder.FolderPath);
                            }

                            for (int i = 1; i <= mi.Attachments.Count; i++)
                            {
                                if (ext.Any(mi.Attachments[i].FileName.ToLower().Contains))
                                {
                                    if (mi.Attachments[i].Type == OlAttachmentType.olByValue)
                                    {
                                        string filename = Path.Combine(basePath + folder.FolderPath, mi.Attachments[i].FileName);
                                        mi.Attachments[i].SaveAsFile(filename);
                                    }

                                }

                            }

                        }
                    }
                }
                catch (System.Exception e)
                {
                    Console.WriteLine(e.Message);
                    Console.WriteLine(e);
                }

            }
        }
    }
}
