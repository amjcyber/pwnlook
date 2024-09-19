using System.Globalization;
using System.Runtime.InteropServices;
using Newtonsoft.Json;
using Redemption;
using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using static System.Collections.Specialized.BitVector32;
using System.Reflection;
using System.ComponentModel;

class Program
{
    static void Main(string[] args)
    {
        // Default var values
        string folderPath = "Inbox";
        DateTime? searchDate = null;
        int latestCount = 0;
        bool listFolders = false;
        string emailId = null;
        bool outputJson = false;
        string searchCriteria = null;
        string searchValue = null;
        int? attachmentIndex = null;
        bool showHelp = false;
        bool listMailboxes = false;
        string mailboxName = null;

        // Parse command-line arguments
        for (int i = 0; i < args.Length; i++)
        {
            switch (args[i])
            {
                case "-folder":
                    if (i + 1 < args.Length)
                    {
                        folderPath = args[i + 1];
                        i++;
                    }
                    break;
                case "-date":
                    if (i + 1 < args.Length && DateTime.TryParseExact(args[i + 1], "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDate))
                    {
                        searchDate = parsedDate;
                        i++;
                    }
                    else
                    {
                        Console.WriteLine($"Invalid date format: {args[i + 1]}. Expected format: YYYY-MM-DD.");
                        return;
                    }
                    break;
                case "-latest":
                    if (i + 1 < args.Length && int.TryParse(args[i + 1], out int parsedCount) && parsedCount > 0)
                    {
                        latestCount = parsedCount;
                        i++;
                    }
                    else
                    {
                        Console.WriteLine($"Invalid latest count: {args[i + 1]}. Expected a positive integer.");
                        return;
                    }
                    break;
                case "-listfolders":
                    listFolders = true;
                    break;
                case "-read":
                    if (i + 1 < args.Length)
                    {
                        emailId = args[i + 1];
                        i++;
                    }
                    break;
                case "-json":
                    outputJson = true;
                    break;
                case "-search":
                    if (i + 1 < args.Length)
                    {
                        searchCriteria = args[i + 1].ToLower();
                        i++;
                    }
                    break;
                case "-value":
                    if (i + 1 < args.Length)
                    {
                        searchValue = args[i + 1];
                        i++;
                    }
                    break;
                case "-attachment":
                    if (i + 1 < args.Length && int.TryParse(args[i + 1], out int index))
                    {
                        attachmentIndex = index;
                        i++;
                    }
                    else
                    {
                        Console.WriteLine($"Invalid attachment index: {args[i + 1]}. Expected an integer.");
                        return;
                    }
                    break;
                case "-h":
                case "--help":
                    showHelp = true;
                    break;
                case "-listmailboxes":
                    listMailboxes = true;
                    break;
                case "-mailbox":
                    if (i + 1 < args.Length)
                    {
                        mailboxName = args[i + 1];
                        i++;
                    }
                    break;
            }
        }

        if (showHelp)
        {
            ShowHelp();
            return;
        }

        RDOSession outlook = null;
        try
        {
            //outlook = new RDOSession();
            outlook = RedemptionLoader.new_RDOSession();
            outlook.Logon(Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            //outlook.Logon();
        }
        catch (Exception ex)
        {
            Console.WriteLine("!!Something went wrong. Maybe Outlook is not running.");
            return;
        }

        try
        {
            if (listMailboxes)
            {
                // List all mailboxes
                ListMailboxes(outlook);
            }
            /*
            else if (!string.IsNullOrEmpty(mailboxName))
            {
                // Set the specified mailbox as the current store
                RDOStore mailboxStore = GetMailboxStore(outlook, mailboxName);
                if (mailboxStore == null)
                {
                    Console.WriteLine($"Mailbox '{mailboxName}' not found.");
                    return;
                }
                //outlook.CurrentStore = mailboxStore;
            }
            */

            else if (!string.IsNullOrEmpty(emailId) && !attachmentIndex.HasValue && mailboxName != null)
            {
                // Fetch and print the email body by ID
                PrintEmailBodyById(outlook, folderPath, emailId, outputJson, mailboxName);
            }
            else if (attachmentIndex.HasValue && !string.IsNullOrEmpty(emailId) && mailboxName != null)
            {
                // Fetch and convert the email attachment to Base64
                PrintAttachmentAsBase64(outlook, folderPath, emailId, attachmentIndex.Value, mailboxName);
            }
            else if (listFolders && mailboxName != null)
            {
                RDOStore selectedStore = GetMailboxByName(outlook, mailboxName);
                // List the entire folder tree
                RDOFolder rootFolder = selectedStore.IPMRootFolder;
                ListFolders(rootFolder, 0);
                ReleaseComObject(rootFolder);
            }
            else if (!string.IsNullOrEmpty(searchCriteria) && !string.IsNullOrEmpty(searchValue) && mailboxName != null)
            {
                // Search for emails
                RDOFolder targetFolder = GetFolderByPath(outlook, folderPath, mailboxName);
                if (targetFolder == null)
                {
                    Console.WriteLine($"Folder '{folderPath}' not found.");
                    return;
                }
                SearchEmails(targetFolder, searchCriteria, searchValue, outputJson);
                ReleaseComObject(targetFolder);
            }
            else if (mailboxName != null)
            {
                // List emails by date or by latest count
                RDOFolder targetFolder = GetFolderByPath(outlook, folderPath, mailboxName);
                if (targetFolder == null)
                {
                    Console.WriteLine($"Folder '{folderPath}' not found.");
                    return;
                }
                if (searchDate.HasValue)
                {
                    ListEmailsFromDate(targetFolder, searchDate.Value, outputJson);
                }
                else if (latestCount > 0)
                {
                    ListLatestEmails(targetFolder, latestCount, outputJson);
                }
                else
                {
                    Console.WriteLine("No valid operation specified. Use the correct arguments.");
                }
            }
            else
            {
                Console.WriteLine("No valid operation specified. Use the correct arguments.");
            }
        }
        finally
        {
            // Logoff
            try
            {
                outlook.Logoff();
                Marshal.ReleaseComObject(outlook);
                ReleaseComObject(outlook);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error logging off: {ex.Message}");
            }
        }
    }

    static RDOStore GetMailboxStore(RDOSession outlook, string mailboxName)
    {
        foreach (RDOStore store in outlook.Stores)
        {
            Console.WriteLine(store.Name);
        }
        return null;
    }

    static RDOStore GetMailboxByName(RDOSession outlook, string mailboxName)
    {
        foreach (RDOStore store in outlook.Stores)
        {
            if (store.Name.Equals(mailboxName, StringComparison.OrdinalIgnoreCase))
            {
                return store;
            }
        }
        return null;
    }
    static RDOFolder GetFolderByPath(RDOSession outlook, string folderPath, string mailboxName)
    {
        RDOStore selectedStore = GetMailboxByName(outlook, mailboxName);
        // List the entire folder tree
        RDOFolder currentFolder = selectedStore.IPMRootFolder;
        //
        //RDOFolder currentFolder = outlook.Stores.DefaultStore.IPMRootFolder;
        string[] pathComponents = folderPath.Split('\\');

        foreach (string folderName in pathComponents)
        {
            currentFolder = FindFolderByName(currentFolder, folderName);
            if (currentFolder == null)
            {
                return null;
            }
        }
        return currentFolder;
    }

    static RDOFolder FindFolderByName(RDOFolder folder, string folderName)
    {
        foreach (RDOFolder subFolder in folder.Folders)
        {
            if (subFolder.Name.Equals(folderName, StringComparison.OrdinalIgnoreCase))
            {
                return subFolder;
            }
        }
        return null;
    }

    static void ListMailboxes(RDOSession outlook)
    {
        RDOStores stores = outlook.Stores;

        Console.WriteLine("Available Mailboxes:");

        foreach (RDOStore store in stores)
        {
            Console.WriteLine($"    - {store.Name}");
        }
    }

    static void ListFolders(RDOFolder folder, int level)
    {
        Console.WriteLine(new string(' ', level * 4) + folder.Name);
        RDOFolders subFolders = folder.Folders;
        foreach (RDOFolder subFolder in subFolders)
        {
            ListFolders(subFolder, level + 1);
        }
    }

    static void ListEmailsFromDate(RDOFolder folder, DateTime searchDate, bool outputJson)
    {
        DateTime startOfDay = searchDate.Date;
        DateTime endOfDay = startOfDay.AddDays(1).AddTicks(-1);

        RDOItems items = folder.Items;

        var emails = items.OfType<RDOMail>()
                          .Where(mail => mail.ReceivedTime >= startOfDay && mail.ReceivedTime <= endOfDay)
                          .OrderByDescending(mail => mail.ReceivedTime)
                          .Select(mail => new EmailDetails
                          {
                              Sender = mail.SenderEmailAddress,
                              // Recipients = string.Join(", ", mail.Recipients.OfType<RDORecipient>().Select(r => r.Address)),
                              Recipients = string.Join(", ", mail.Recipients.OfType<RDORecipient>().Select(r => r.Address).ToArray()),
                              Subject = mail.Subject,
                              Attachments = mail.Attachments.OfType<RDOAttachment>().Select(a => a.FileName).ToList(),
                              Date = mail.ReceivedTime.ToString("yyyy-MM-dd"),
                              Folder = folder.Name,
                              ID = mail.EntryID
                          })
                          .ToList();

        if (outputJson)
        {
            Console.WriteLine(JsonConvert.SerializeObject(emails, Formatting.Indented));
        }
        else
        {
            foreach (var email in emails)
            {
                //Console.WriteLine($"Subject: {email.Subject}, Date: {email.Date}, Sender: {email.Sender}, Recipients: {email.Recipients}, Attachments: {string.Join(", ", email.Attachments)}, ID: {email.ID}");
                Console.WriteLine($"Subject: {email.Subject}, Date: {email.Date}, Sender: {email.Sender}, Recipients: {email.Recipients}, Attachments: {string.Join(", ", email.Attachments.ToArray())}, ID: {email.ID}");
            }
        }
    }

    static void ListLatestEmails(RDOFolder folder, int latestCount, bool outputJson)
    {
        RDOItems items = folder.Items;
        var latestEmails = items.OfType<RDOMail>()
                                .OrderByDescending(mail => mail.ReceivedTime)
                                .Take(latestCount)
                                .Select(mail => new EmailDetails
                                {
                                    Sender = mail.SenderEmailAddress,
                                    // Recipients = string.Join(", ", mail.Recipients.OfType<RDORecipient>().Select(r => r.Address)),
                                    Recipients = string.Join(", ", mail.Recipients.OfType<RDORecipient>().Select(r => r.Address).ToArray()),
                                    Subject = mail.Subject,
                                    Attachments = mail.Attachments.OfType<RDOAttachment>().Select(a => a.FileName).ToList(),
                                    Date = mail.ReceivedTime.ToString("yyyy-MM-dd"),
                                    Folder = folder.Name,
                                    ID = mail.EntryID
                                })
                                .ToList();

        if (outputJson)
        {
            Console.WriteLine(JsonConvert.SerializeObject(latestEmails, Formatting.Indented));
        }
        else
        {
            foreach (var email in latestEmails)
            {
                //Console.WriteLine($"Subject: {email.Subject}, Date: {email.Date}, Sender: {email.Sender}, Recipients: {email.Recipients}, Attachments: {string.Join(", ", email.Attachments)}, ID: {email.ID}");
                Console.WriteLine($"Subject: {email.Subject}, Date: {email.Date}, Sender: {email.Sender}, Recipients: {email.Recipients}, Attachments: {string.Join(", ", email.Attachments.ToArray())}, ID: {email.ID}");
            }
        }
    }

    static void PrintEmailBodyById(RDOSession outlook, string folderPath, string emailId, bool outputJson, string mailboxName)
    {
        RDOFolder targetFolder = GetFolderByPath(outlook, folderPath, mailboxName);
        if (targetFolder == null)
        {
            Console.WriteLine($"Folder '{folderPath}' not found.");
            return;
        }

        RDOMail email = targetFolder.Items.Find($"[EntryID]='{emailId}'") as RDOMail;
        if (email != null)
        {
            if (outputJson)
            {
                var emailDetails = new EmailDetails
                {
                    Sender = email.SenderEmailAddress,
                    // Recipients = string.Join(", ", email.Recipients.OfType<RDORecipient>().Select(r => r.Address)),
                    Recipients = string.Join(", ", email.Recipients.OfType<RDORecipient>().Select(r => r.Address).ToArray()),
                    Subject = email.Subject,
                    Body = email.Body,
                    Attachments = email.Attachments.OfType<RDOAttachment>().Select(a => a.FileName).ToList(),
                    Date = email.ReceivedTime.ToString("yyyy-MM-dd"),
                    Folder = targetFolder.Name,
                    ID = email.EntryID
                };
                Console.WriteLine(JsonConvert.SerializeObject(emailDetails, Formatting.Indented));
            }
            else
            {
                Console.WriteLine($"Subject: {email.Subject}, Date: {email.ReceivedTime}, Sender: {email.SenderEmailAddress}, Recipients: {string.Join(", ", email.Recipients.OfType<RDORecipient>().Select(r => r.Address).ToArray())}");
                Console.WriteLine($"Body: {email.Body}");
            }
        }
        else
        {
            Console.WriteLine($"Email with ID '{emailId}' not found.");
        }
    }

    static void PrintAttachmentAsBase64(RDOSession outlook, string folderPath, string emailId, int attachmentIndex, string mailboxName)
    {
        RDOFolder targetFolder = GetFolderByPath(outlook, folderPath, mailboxName);
        if (targetFolder == null)
        {
            Console.WriteLine($"Folder '{folderPath}' not found.");
            return;
        }

        // Find the email by ID
        RDOMail mail = targetFolder.Items.OfType<RDOMail>().FirstOrDefault(m => m.EntryID == emailId);
        if (mail == null)
        {
            Console.WriteLine($"Email with ID '{emailId}' not found in folder '{folderPath}'.");
            return;
        }

        // Get the attachment
        RDOAttachment attachment = mail.Attachments.OfType<RDOAttachment>().ElementAtOrDefault(attachmentIndex);
        if (attachment == null)
        {
            Console.WriteLine($"Attachment index {attachmentIndex} not found in email with ID '{emailId}'.");
            return;
        }

        // Save attachment to a temporary file
        string tempFilePath = Path.GetTempFileName();
        try
        {
            attachment.SaveAsFile(tempFilePath);
            byte[] fileBytes = File.ReadAllBytes(tempFilePath);
            string base64String = Convert.ToBase64String(fileBytes);

            // Print the Base64 string
            Console.WriteLine(base64String);
        }
        finally
        {
            // Delete the temporary file
            if (File.Exists(tempFilePath))
            {
                File.Delete(tempFilePath);
            }
        }
    }

    static void SearchEmails(RDOFolder folder, string searchCriteria, string searchValue, bool outputJson)
    {
        RDOItems items = folder.Items;
        var matchingEmails = items.OfType<RDOMail>()
                                  .Where(mail =>
                                  {
                                      switch (searchCriteria)
                                      {
                                          case "subject":
                                              //return mail.Subject.Contains(searchValue, StringComparison.OrdinalIgnoreCase);
                                              return mail.Subject.IndexOf(searchValue, StringComparison.OrdinalIgnoreCase) >= 0;
                                          case "sender":
                                              //return mail.SenderEmailAddress.Contains(searchValue, StringComparison.OrdinalIgnoreCase);
                                              return mail.SenderEmailAddress.IndexOf(searchValue, StringComparison.OrdinalIgnoreCase) >= 0;
                                          case "recipient":
                                              //return mail.Recipients.OfType<RDORecipient>().Any(r => r.Address.Contains(searchValue, StringComparison.OrdinalIgnoreCase));
                                              //return mail.Recipients.IndexOf(searchValue, StringComparison.OrdinalIgnoreCase) >= 0;
                                              return mail.Recipients.OfType<RDORecipient>().Any(r => r.Address.IndexOf(searchValue, StringComparison.OrdinalIgnoreCase) >= 0);
                                          case "body":
                                              //return mail.Body.Contains(searchValue, StringComparison.OrdinalIgnoreCase);
                                              return mail.Body.IndexOf(searchValue, StringComparison.OrdinalIgnoreCase) >= 0;
                                          default:
                                              return false;
                                      }
                                  })
                                  .OrderByDescending(mail => mail.ReceivedTime)
                                  .Select(mail => new EmailDetails
                                  {
                                      Sender = mail.SenderEmailAddress,
                                      Recipients = string.Join(", ", mail.Recipients.OfType<RDORecipient>().Select(r => r.Address).ToArray()),
                                      Subject = mail.Subject,
                                      Attachments = mail.Attachments.OfType<RDOAttachment>().Select(a => a.FileName).ToList(),
                                      Date = mail.ReceivedTime.ToString("yyyy-MM-dd"),
                                      Folder = folder.Name,
                                      ID = mail.EntryID
                                  })
                                  .ToList();

        if (outputJson)
        {
            Console.WriteLine(JsonConvert.SerializeObject(matchingEmails, Formatting.Indented));
        }
        else
        {
            foreach (var email in matchingEmails)
            {
                //Console.WriteLine($"Subject: {email.Subject}, Date: {email.Date}, Sender: {email.Sender}, Recipients: {email.Recipients}, Attachments: {string.Join(", ", email.Attachments)}, ID: {email.ID}");
                Console.WriteLine($"Subject: {email.Subject}, Date: {email.Date}, Sender: {email.Sender}, Recipients: {email.Recipients}, Attachments: {string.Join(", ", email.Attachments.ToArray())}, ID: {email.ID}");
            }
        }
    }

    static void ShowHelp()
    {
        Console.WriteLine("\r\n\r\n                    .__                 __    \r\n________  _  ______ |  |   ____   ____ |  | __\r\n\\____ \\ \\/ \\/ /    \\|  |  /  _ \\ /  _ \\|  |/ /\r\n|  |_> >     /   |  \\  |_(  <_> |  <_> )    < \r\n|   __/ \\/\\_/|___|  /____/\\____/ \\____/|__|_ \\\r\n|__|              \\/                        \\/\r\n\r\n");
        Console.WriteLine("Usage: pwnlook.exe [options]");
        Console.WriteLine();
        Console.WriteLine("List mailboxes:");
        Console.WriteLine("  -listmailboxes");
        Console.WriteLine();
        Console.WriteLine("List folders:");
        Console.WriteLine("  -mailbox <mailbox> -listfolders");
        Console.WriteLine();
        Console.WriteLine("List emails from date:");
        Console.WriteLine("  -mailbox <mailbox> -folder <Folder\\Path> -date <yyyy-MM-dd>");
        Console.WriteLine();
        Console.WriteLine("List latest X emails from folder:");
        Console.WriteLine("  -mailbox <mailbox> -folder <Folder\\Path> -latest <X>");
        Console.WriteLine();
        Console.WriteLine("Read email:");
        Console.WriteLine("  -mailbox <mailbox> -folder <Folder\\Path> -id <ID>");
        Console.WriteLine();
        Console.WriteLine("Download attachment (base64):");
        Console.WriteLine("  -mailbox <mailbox> -folder <Folder\\Path> -id <ID> -attachment <X>");
        Console.WriteLine();
        Console.WriteLine("Search by sender or subject:");
        Console.WriteLine("  -mailbox <mailbox> -folder <Folder\\Path> -search <sender|subject> -value <string>");
        Console.WriteLine();
        Console.WriteLine("Result format in JSON");
        Console.WriteLine("  -json");
        Console.WriteLine();
        Console.WriteLine("Examples:");
        Console.WriteLine(".\\pwnlook.exe -mailbox my@mail.com -folder \"Inbox\" -latest 20 -json        Lists latest 20 emails from Inbox");
        Console.WriteLine();
    }

    static void ReleaseComObject(object obj)
    {
        if (obj != null && Marshal.IsComObject(obj))
        {
            Marshal.ReleaseComObject(obj);
        }
    }
}

// Email details class to hold the extracted data
class EmailDetails
{
    public string Sender { get; set; }
    public string Recipients { get; set; }
    public string Subject { get; set; }
    public string Body { get; set; }
    public List<string> Attachments { get; set; }
    public string Date { get; set; }
    public string Folder { get; set; }
    public string ID { get; set; }
}

namespace Redemption
{
    public static class RedemptionLoader
    {
        #region public methods
        //64 bit dll location - defaults to <assemblydir>\Redemption64.dll
        public static string DllLocation64Bit;
        //32 bit dll location - defaults to <assemblydir>\Redemption.dll
        public static string DllLocation32Bit;


        //The only creatable RDO object - RDOSession
        public static RDOSession new_RDOSession()
        {
            return (RDOSession)NewRedemptionObject(new Guid("29AB7A12-B531-450E-8F7A-EA94C2F3C05F"));
        }

        //Safe*Item objects
        public static SafeMailItem new_SafeMailItem()
        {
            return (SafeMailItem)NewRedemptionObject(new Guid("741BEEFD-AEC0-4AFF-84AF-4F61D15F5526"));
        }

        public static SafeContactItem new_SafeContactItem()
        {
            return (SafeContactItem)NewRedemptionObject(new Guid("4FD5C4D3-6C15-4EA0-9EB9-EEE8FC74A91B"));
        }

        public static SafeAppointmentItem new_SafeAppointmentItem()
        {
            return (SafeAppointmentItem)NewRedemptionObject(new Guid("620D55B0-F2FB-464E-A278-B4308DB1DB2B"));
        }

        public static SafeTaskItem new_SafeTaskItem()
        {
            return (SafeTaskItem)NewRedemptionObject(new Guid("7A41359E-0407-470F-B3F7-7C6A0F7C449A"));
        }

        public static SafeJournalItem new_SafeJournalItem()
        {
            return (SafeJournalItem)NewRedemptionObject(new Guid("C5AA36A1-8BD1-47E0-90F8-47E7239C6EA1"));
        }

        public static SafeMeetingItem new_SafeMeetingItem()
        {
            return (SafeMeetingItem)NewRedemptionObject(new Guid("FA2CBAFB-F7B1-4F41-9B7A-73329A6C1CB7"));
        }

        public static SafePostItem new_SafePostItem()
        {
            return (SafePostItem)NewRedemptionObject(new Guid("11E2BC0C-5D4F-4E0C-B438-501FFE05A382"));
        }

        public static SafeReportItem new_SafeReportItem()
        {
            return (SafeReportItem)NewRedemptionObject(new Guid("D46BA7B2-899F-4F60-85C7-4DF5713F6F18"));
        }

        public static MAPIFolder new_MAPIFolder()
        {
            return (MAPIFolder)NewRedemptionObject(new Guid("03C4C5F4-1893-444C-B8D8-002F0034DA92"));
        }

        public static SafeCurrentUser new_SafeCurrentUser()
        {
            return (SafeCurrentUser)NewRedemptionObject(new Guid("7ED1E9B1-CB57-4FA0-84E8-FAE653FE8E6B"));
        }

        public static SafeDistList new_SafeDistList()
        {
            return (SafeDistList)NewRedemptionObject(new Guid("7C4A630A-DE98-4E3E-8093-E8F5E159BB72"));
        }

        public static AddressLists new_AddressLists()
        {
            return (AddressLists)NewRedemptionObject(new Guid("37587889-FC28-4507-B6D3-8557305F7511"));
        }

        public static MAPITable new_MAPITable()
        {
            return (MAPITable)NewRedemptionObject(new Guid("A6931B16-90FA-4D69-A49F-3ABFA2C04060"));
        }

        public static MAPIUtils new_MAPIUtils()
        {
            return (MAPIUtils)NewRedemptionObject(new Guid("4A5E947E-C407-4DCC-A0B5-5658E457153B"));
        }

        public static SafeInspector new_SafeInspector()
        {
            return (SafeInspector)NewRedemptionObject(new Guid("ED323630-B4FD-4628-BC6A-D4CC44AE3F00"));
        }

        public static SafeExplorer new_SafeExplorer()
        {
            return (SafeExplorer)NewRedemptionObject(new Guid("C3B05695-AE2C-4FD5-A191-2E4C782C03E0"));
        }

        public static SafeApplication new_SafeApplication()
        {
            return (SafeApplication)NewRedemptionObject(new Guid("9DCB6F1D-9AB2-4002-A469-89A940E28A75"));
        }

        #endregion


        #region private methods



        static RedemptionLoader()
        {
            //default locations of the dlls

            //use CodeBase instead of Location because of Shadow Copy.
            string codebase = Assembly.GetExecutingAssembly().CodeBase;
            var vUri = new UriBuilder(codebase);
            string vPath = Uri.UnescapeDataString(vUri.Path + vUri.Fragment);
            string directory = Path.GetDirectoryName(vPath);
            if (!string.IsNullOrEmpty(vUri.Host)) directory = @"\\" + vUri.Host + directory;
            DllLocation64Bit = Path.Combine(directory, "redemption64.dll");
            DllLocation32Bit = Path.Combine(directory, "redemption.dll");
        }

        [ComVisible(false)]
        [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("00000001-0000-0000-C000-000000000046")]
        private interface IClassFactory
        {
            void CreateInstance([MarshalAs(UnmanagedType.Interface)] object pUnkOuter, ref Guid refiid, [MarshalAs(UnmanagedType.Interface)] out object ppunk);
            void LockServer(bool fLock);
        }

        [ComVisible(false)]
        [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("00000000-0000-0000-C000-000000000046")]
        private interface IUnknown
        {
        }

        private delegate int DllGetClassObject(ref Guid ClassId, ref Guid InterfaceId, [Out, MarshalAs(UnmanagedType.Interface)] out object ppunk);
        private delegate int DllCanUnloadNow();

        //COM GUIDs
        private static Guid IID_IClassFactory = new Guid("00000001-0000-0000-C000-000000000046");
        private static Guid IID_IUnknown = new Guid("00000000-0000-0000-C000-000000000046");

        //win32 functions to load\unload dlls and get a function pointer 
        private class Win32NativeMethods
        {
            [DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true)]
            public static extern IntPtr GetProcAddress(IntPtr hModule, string lpProcName);

            [DllImport("kernel32.dll", SetLastError = true)]
            public static extern bool FreeLibrary(IntPtr hModule);

            [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
            public static extern IntPtr LoadLibraryW(string lpFileName);
        }

        //private variables
        private static IntPtr _redemptionDllHandle = IntPtr.Zero;
        private static IntPtr _dllGetClassObjectPtr = IntPtr.Zero;
        private static DllGetClassObject _dllGetClassObject;
        private static readonly object _criticalSection = new object();

        private static IUnknown NewRedemptionObject(Guid guid)
        {
            object res = null;

            lock (_criticalSection)
            {
                IClassFactory ClassFactory;
                if (_redemptionDllHandle.Equals(IntPtr.Zero))
                {
                    string dllPath;
                    if (IntPtr.Size == 8) dllPath = DllLocation64Bit;
                    else dllPath = DllLocation32Bit;
                    _redemptionDllHandle = Win32NativeMethods.LoadLibraryW(dllPath);
                    if (_redemptionDllHandle.Equals(IntPtr.Zero))
                        //throw new Exception(string.Format("Could not load '{0}'\nMake sure the dll exists.", dllPath));
                        throw new Win32Exception(Marshal.GetLastWin32Error());
                    _dllGetClassObjectPtr = Win32NativeMethods.GetProcAddress(_redemptionDllHandle, "DllGetClassObject");
                    if (_dllGetClassObjectPtr.Equals(IntPtr.Zero))
                        //throw new Exception("Could not retrieve a pointer to the 'DllGetClassObject' function exported by the dll");
                        throw new Win32Exception(Marshal.GetLastWin32Error());
                    _dllGetClassObject =
                        (DllGetClassObject)
                        Marshal.GetDelegateForFunctionPointer(_dllGetClassObjectPtr, typeof(DllGetClassObject));
                }

                Object unk;
                int hr = _dllGetClassObject(ref guid, ref IID_IClassFactory, out unk);
                if (hr != 0) throw new Exception("DllGetClassObject failed with error code 0x" + hr.ToString("x8"));
                ClassFactory = unk as IClassFactory;
                ClassFactory.CreateInstance(null, ref IID_IUnknown, out res);

                Marshal.ReleaseComObject(unk);
                Marshal.ReleaseComObject(ClassFactory);
            } //lock

            return (res as IUnknown);
        }

        #endregion

    }
}