using System;
using System.Linq;
using System.IO;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ConsoleApp1
{
    class Program
    {

        // Path where attachments will be saved
        static string basePath = @"c:\temp\emails\";
        static int totalfilesize = 0;

        static void Main(string[] args)
        {
            //EnumerateAccounts();
            EnumerateFoldersInDefaultStore();
            //Console.WriteLine("Total file size:" + totalfilesize);
        }

        static void EnumerateFoldersInDefaultStore()
        {
            Outlook.NameSpace ns = null;
            Outlook.Stores stores = null;
            Outlook.Folder rootFolder = null;
            string storePath = string.Empty;

            storePath = @"C:\temp\emails\archive.pst";
            

            Outlook.Application Application = new Outlook.Application();
            ns = Application.Session;
            ns.AddStore(storePath);
            stores = ns.Stores;

            var root = stores[2].GetRootFolder() as Outlook.Folder;
            EnumerateFolders(root);
        }

        // Uses recursion to enumerate Outlook subfolders.
        static void EnumerateFolders(Outlook.Folder folder)
        {
            Outlook.Folders childFolders = folder.Folders;
            if (childFolders.Count > 0)
            {
                foreach (Outlook.Folder childFolder in childFolders)
                {
                    //if (childFolder.FolderPath.Contains("Inbox"))
                    if (childFolder.FolderPath.Contains("Sent Items"))
                        {
                            //Console.WriteLine(childFolder.FolderPath);
                            // Call EnumerateFolders using childFolder, to see if there are any sub-folders within this one
                            //EnumerateFolders(childFolder);
                            Console.WriteLine("Checking in " + childFolder.FolderPath);
                        IterateMessages(childFolder);
                    }
                }
            }

            //Console.WriteLine("Checking in " + folder.FolderPath);
            //IterateMessages(folder);

        }

        static void IterateMessages(Outlook.Folder folder)
        {
            // Iterate through all items ("messages") in a folder
            Outlook.Items fi = folder.Items;
            if (fi != null)
            {
                foreach (object item in fi)
                {
                    try
                    {
                        Outlook.MailItem mi = (Outlook.MailItem)item;
                        var attachments = mi.Attachments;
                        if (attachments.Count != 0)
                        {
                            var senderName = mi.SenderName.Replace(":", "").Replace("/", "").Replace("\\", "").Replace("*", "").Replace("*", "").Replace("<", "").Replace(">", "").Replace("|", "").Replace("\"", "").Replace("'", "");
                            senderName = senderName.Trim();


                            var dirPath = basePath + folder.FolderPath + "\\" + senderName;
                            // Create a directory to store the attachment 
                            if (!Directory.Exists(dirPath))
                            {
                                Directory.CreateDirectory(dirPath);
                            }

                            if (mi.Subject != null)
                            {
                                var subject = mi.Subject.Replace(":", "").Replace("/", "").Replace("\\", "").Replace("*", "").Replace("*", "").Replace("<", "").Replace(">", "").Replace("|", "").Replace("\"", "").Replace("'", "");
                                subject = subject.Trim();

                                Console.WriteLine("\n" + dirPath + "\\" + subject);
                                for (int i = 1; i <= mi.Attachments.Count; i++)
                                {
                                    // var fn = mi.Attachments[i].FileName.ToLower();
                                    // Create a further sub-folder for the sender

                                    var filename = mi.Attachments[i].FileName;
                                    if (!Directory.Exists(dirPath + "\\" + subject))
                                    {
                                        Directory.CreateDirectory(dirPath + "\\" + subject);
                                    }
                                    // totalfilesize = totalfilesize + mi.Attachments[i].Size;
                                    
                                    if (!File.Exists(dirPath + "\\" + subject + "\\" + filename))
                                    {
                                        Console.WriteLine("Saving: " + dirPath + "\\" + subject + "\\" + filename);
                                        mi.Attachments[i].SaveAsFile(dirPath + "\\" + subject + "\\" + filename);
                                        
                                        //mi.Attachments[i].Delete();
                                    }
                                    else
                                    {
                                        //Console.WriteLine("Already saved " + mi.Attachments[i].FileName);
                                    }
                                }
                            }
                        }
                    }

                    catch (Exception e)
                    {
                        Console.WriteLine("An error occurred: '{0}'", e);
                    }
                }
            }
        }

        // Retrieves the email address for a given account object
        static string EnumerateAccountEmailAddress(Outlook.Account account)
        {
            try
            {
                if (string.IsNullOrEmpty(account.SmtpAddress) || string.IsNullOrEmpty(account.UserName))
                {
                    Outlook.AddressEntry oAE = account.CurrentUser.AddressEntry as Outlook.AddressEntry;
                    if (oAE.Type == "EX")
                    {
                        Outlook.ExchangeUser oEU = oAE.GetExchangeUser() as Outlook.ExchangeUser;
                        return oEU.PrimarySmtpAddress;
                    }
                    else
                    {
                        return oAE.Address;
                    }
                }
                else
                {
                    return account.SmtpAddress;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return "";
            }
        }

        static void EnumerateAccounts()
        {
            Console.Clear();
            Console.WriteLine("Outlook Attachment Extractor v0.1");
            Console.WriteLine("---------------------------------");
            int id;
            Outlook.Application Application = new Outlook.Application();
            Outlook.Accounts accounts = Application.Session.Accounts;

            string response = "";
            while (true == true)
            {

                id = 1;
                foreach (Outlook.Account account in accounts)
                {
                    Console.WriteLine(id + ":" + EnumerateAccountEmailAddress(account));
                    id++;
                }
                Console.WriteLine("Q: Quit Application");

                response = Console.ReadLine().ToUpper();
                if (response == "Q")
                {
                    Console.WriteLine("Quitting");
                    return;
                }
                if (response != "")
                {
                    if (Int32.Parse(response.Trim()) >= 1 && Int32.Parse(response.Trim()) < id)
                    {
                        Console.WriteLine("Processing: " + accounts[Int32.Parse(response.Trim())].DisplayName);
                        Console.WriteLine("Processing: " + EnumerateAccountEmailAddress(accounts[Int32.Parse(response.Trim())]));

                        Outlook.Folder selectedFolder = Application.Session.DefaultStore.GetRootFolder() as Outlook.Folder;
                        selectedFolder = GetFolder(@"\\" + accounts[Int32.Parse(response.Trim())].DisplayName);
                        EnumerateFolders(selectedFolder);
                        Console.WriteLine("Finished Processing " + accounts[Int32.Parse(response.Trim())].DisplayName);
                        Console.WriteLine("");
                    }
                    else
                    {
                        Console.WriteLine("Invalid Account Selected");
                    }
                }
            }

        }

        // Returns Folder object based on folder path
        static Outlook.Folder GetFolder(string folderPath)
        {
            Console.WriteLine("Looking for: " + folderPath);
            Outlook.Folder folder;
            string backslash = @"\";
            try
            {
                if (folderPath.StartsWith(@"\\"))
                {
                    folderPath = folderPath.Remove(0, 2);
                }
                String[] folders = folderPath.Split(backslash.ToCharArray());
                Outlook.Application Application = new Outlook.Application();
                folder = Application.Session.Folders[folders[0]] as Outlook.Folder;
                if (folder != null)
                {
                    for (int i = 1; i <= folders.GetUpperBound(0); i++)
                    {
                        Outlook.Folders subFolders = folder.Folders;
                        folder = subFolders[folders[i]] as Outlook.Folder;
                        if (folder == null)
                        {
                            return null;
                        }
                    }
                }
                return folder;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return null;
            }
        }

    }


}