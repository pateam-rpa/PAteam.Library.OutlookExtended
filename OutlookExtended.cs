using Direct.Interface;
using Direct.Shared;
using System.Collections.Generic;
using Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;
using System;
using Direct.Office.Library;
using System.Threading;
using System.Diagnostics;

namespace PAteam.Library.OutlookExtended
{
    [DirectSealed]
    [ParameterType(false)]
    [DirectDom("Microsoft Outlook Extended", "Microsoft Outlook Extended")]
    public static class OutlookExtended
    {
        private static IDirectLog logger = DirectLogManager.GetLogger("LibraryObjects");
        private static DateTime classInitializedTime = DateTime.Now;
        private static string outlookNotInstalledErrorMsg = "Outlook may not be installed or COM object cannot be registered.";
        private static bool IsOutlookInstalled = false;

        private static Application Application { get; set; }

        static OutlookExtended()
        {
            TryInitializeOutlookApplication();

            if (!IsOutlookInstalled)
            {
                logger.Error(outlookNotInstalledErrorMsg);
                throw new System.Exception(outlookNotInstalledErrorMsg);
            }
        }

        [DirectDom("Get Report Emails")]
        [DefaultParameter("top", 10)]
        [DefaultParameter("onlyUnread", false)]
        [DefaultParameter("markAsRead", false)]
        [DirectDomMethod("Get Report Emails: Max Emails to Get {Top}, Folder {Folder}, Only Unread {Only Unread}, Mark as Read {Mark as Read}, From Date {From Date}, To Date {To Date}")]
        [MethodDescription("Get maximum 'Top' number of report emails from folder 'Folder', specifying weather you want only unread emails and/or mark retrieved emails as read.")]
        public static DirectCollection<OutlookReportEmail> GetReportEmails(
            int top,
            string folderPath,
            bool onlyUnread,
            bool markAsRead,
            DateTime fromDate,
            DateTime toDate)
        {
            //add new param to only query top x
            logger.InfoFormat("getting emails: max {0} emails, from folder '{1}', unread only - [{2}], mark as read = [{3}]", top, folderPath, onlyUnread, markAsRead);
            DirectCollection<OutlookReportEmail> emails = new DirectCollection<OutlookReportEmail>();

            if (IsOutlookRestarted())
            {
                TryInitializeOutlookApplication();
            }

            if (IsOutlookInstalled)
            {
                try
                {
                    MAPIFolder folderFromFolderPath = GetFolderFromFolderPath(folderPath);
                    if (folderFromFolderPath == null)
                    {
                        logger.Error("Folder not found");
                        return new DirectCollection<OutlookReportEmail>();
                    }
                    Items allItems = folderFromFolderPath.Items;
                    Items outlookItems = onlyUnread ? allItems.Restrict("[Unread]=true") : allItems;

                    if (fromDate != DateTime.MinValue)
                    {
                        outlookItems = outlookItems.Restrict(string.Format("[ReceivedTime]>='{0}'", fromDate.ToString("g")));
                    }

                    if (toDate != DateTime.MinValue)
                    {
                        outlookItems = outlookItems.Restrict(string.Format("[ReceivedTime]<='{0}'", toDate.ToString("g")));
                    }

                    outlookItems.Sort("[ReceivedTime]", true);
                    List<OutlookReportEmail> mails = GetOutlookReportEmailsFromItems(outlookItems, top);

                    if (markAsRead)
                    {
                        MarkEmailsAsRead(mails);
                    }

                    emails = mails;

                }
                catch (System.Exception ex)
                {
                    logger.Error("Unhandled exception happened: ", ex);
                }

            }
            else
            {
                logger.Error(outlookNotInstalledErrorMsg);
            }

            return emails;

        }

        [DirectDom("Move Email to Folder")]
        [DirectDomMethod("Move Email as Read: Email Id {Email Id}, New Folder Path {Folder}")]
        [MethodDescription("Move email with given id to the specified folder. Email ids are assigned by Outlook application.")]
        public static bool MoveEmailToFolder(string id, string emailFolder)
        {
            if (IsOutlookRestarted())
            {
                TryInitializeOutlookApplication();
            }

            if (!IsOutlookInstalled)
            {
                logger.Error(outlookNotInstalledErrorMsg);
                return false;
            }
            else
            {
                try
                {
                    NameSpace mapiNamespace = Application.GetNamespace("MAPI");
                    object outlookItem = mapiNamespace.GetItemFromID(id);


                    MailItem mailItemFromId = outlookItem as MailItem;

                    if (mailItemFromId != null)
                    {
                        MAPIFolder folderFromFolderPath = GetFolderFromFolderPath(emailFolder);
                        mailItemFromId.Move(folderFromFolderPath);
                        return true;
                    }

                    ReportItem reportItemFromId = outlookItem as ReportItem;

                    if (reportItemFromId != null)
                    {
                        MAPIFolder folderFromFolderPath = GetFolderFromFolderPath(emailFolder);
                        reportItemFromId.Move(folderFromFolderPath);
                        return true;
                    }

                    logger.ErrorFormat("no mail item with id {0} was found", id);
                    return false;

                }
                catch (System.Exception ex)
                {
                    logger.ErrorFormat("MoveEmailToFolder() Exception - call stack:\n\t'{0}'\n\t message: '{1}'", ex.StackTrace, ex.Message);
                    return false;
                }
            }
        }

        [DirectDom("Mark Email as Read/Unread")]
        [DirectDomMethod("Mark Email as Read/Unread {Email Status}: Email Id {Email Id}")]
        [MethodDescription("Mark email with given id as read/unread. Email ids are assigned by Outlook application.")]
        public static bool MarkEmailAsRead(bool status, string id)
        {
            if (IsOutlookRestarted())
            {
                TryInitializeOutlookApplication();
            }

            if (!IsOutlookInstalled)
            {
                logger.Error(outlookNotInstalledErrorMsg);
                return false;
            }
            else
            {
                try
                {
                    NameSpace mapiNamespace = Application.GetNamespace("MAPI");
                    object outlookItem = mapiNamespace.GetItemFromID(id);


                    MailItem mailItemFromId = outlookItem as MailItem;

                    if (mailItemFromId != null)
                    {
                        mailItemFromId.UnRead = status;
                        return true;
                    }

                    ReportItem reportItemFromId = outlookItem as ReportItem;

                    if (reportItemFromId != null)
                    {
                        reportItemFromId.UnRead = status;
                        return true;
                    }

                    logger.ErrorFormat("no mail item with id {0} was found", id);
                    return false;

                }
                catch (System.Exception ex)
                {
                    logger.ErrorFormat("MoveEmailToFolder() Exception - call stack:\n\t'{0}'\n\t message: '{1}'", ex.StackTrace, ex.Message);
                    return false;
                }
            }
        }

        [DirectDom("Get Sender Name from Mail Item")]
        [DirectDomMethod("Get Sender Name from Mail Item with Id {Email Id}")]
        [MethodDescription("Returns sender name from mail based on id")]
        public static string GetSenderName(string id)
        {
            if (IsOutlookRestarted())
            {
                TryInitializeOutlookApplication();
            }

            if (!IsOutlookInstalled)
            {
                logger.Error(outlookNotInstalledErrorMsg);
                return "";
            }
            else
            {
                try
                {
                    NameSpace mapiNamespace = Application.GetNamespace("MAPI");
                    object outlookItem = mapiNamespace.GetItemFromID(id);

                    MailItem mailItemFromId = outlookItem as MailItem;

                    if (mailItemFromId != null)
                    {
                        return mailItemFromId.SenderName;
                    }

                    logger.ErrorFormat("GetSenderName() - No mail item with id {0} was found", id);
                    return "";

                }
                catch (System.Exception ex)
                {
                    logger.ErrorFormat("GetSenderName() Exception - call stack:\n\t'{0}'\n\t message: '{1}'", ex.StackTrace, ex.Message);
                    return "";
                }
            }
        }

        [DirectDom("Count Mail Items")]
        [DirectDomMethod("Count mail items from folder {Folder}, with optional subject {Subject}, only unread {Only Unread}")]
        [MethodDescription("Will return a total count of mail items from given folder. You can supply subject if you wish.")]
        public static int CountMailItems(string folderPath, string subject, bool onlyUnread)
        {
            if (IsOutlookRestarted())
            {
                TryInitializeOutlookApplication();
            }

            try
            {
                if (IsOutlookInstalled)
                {

                    MAPIFolder folderFromFolderPath = GetFolderFromFolderPath(folderPath);
                    if (folderFromFolderPath == null)
                    {
                        logger.Error("Folder not found");
                        return 0;
                    }
                    Items allItems = folderFromFolderPath.Items;
                    Items outlookItems = onlyUnread ? allItems.Restrict("[Unread]=true") : allItems;
                    if (!string.IsNullOrEmpty(subject))
                    {
                        outlookItems = outlookItems.Restrict(string.Format("@SQL=\"urn:schemas:httpmail:subject\" like '%{0}%'", subject));
                    }

                    return outlookItems.Count;
                }
                else
                {
                    logger.Error(outlookNotInstalledErrorMsg);
                }
            }
            catch (System.Exception ex)
            {
                logger.Error("CountMailItems(): unexpected error happened: ", ex);
            }


            return 0;
        }

        [DirectDom("Count Report Items")]
        [DirectDomMethod("Count report items from folder {Folder}, only unread {Only Unread}")]
        [MethodDescription("Will return a total count of report items from given folder.")]
        public static int CountReportItems(string folderPath, bool onlyUnread)
        {
            if (IsOutlookRestarted())
            {
                TryInitializeOutlookApplication();
            }

            try
            {
                if (IsOutlookInstalled)
                {

                    MAPIFolder folderFromFolderPath = GetFolderFromFolderPath(folderPath);
                    if (folderFromFolderPath == null)
                    {
                        logger.Error("Folder not found");
                        return 0;
                    }
                    Items allItems = folderFromFolderPath.Items;
                    Items outlookItems = onlyUnread ? allItems.Restrict("[Unread]=true") : allItems;

                    List<OutlookReportEmail> reportEmails = GetOutlookReportEmailsFromItems(outlookItems);

                    return reportEmails.Count;
                }
                else
                {
                    logger.Error(outlookNotInstalledErrorMsg);
                }

            }
            catch (System.Exception ex)
            {
                logger.Error("CountReportItems(): unexpected error happened: ", ex);
            }

            return 0;
        }


        [DirectDom("Get Most Recent Report Item")]
        [DirectDomMethod("Get most recent report item from folder {Folder}, only unread {Only Unread}")]
        [MethodDescription("Will return a most recent report item from given folder")]
        public static OutlookReportEmail GetReportItem(string folderPath, bool onlyUnread)
        {
            if (IsOutlookRestarted())
            {
                TryInitializeOutlookApplication();
            }

            try
            {
                if (IsOutlookInstalled)
                {

                    MAPIFolder folderFromFolderPath = GetFolderFromFolderPath(folderPath);
                    if (folderFromFolderPath == null)
                    {
                        logger.Error("Folder not found");
                        return new OutlookReportEmail();
                    }
                    Items allItems = folderFromFolderPath.Items;
                    Items outlookItems = onlyUnread ? allItems.Restrict("[Unread]=true") : allItems;

                    outlookItems.Sort("[ReceivedTime]", true);

                    return GetOutlookReportEmailsFromItems(outlookItems, 1)[0];
                }
                else
                {
                    logger.Error(outlookNotInstalledErrorMsg);
                }

            }
            catch (System.Exception ex)
            {
                logger.Error("GetReportItem(): unexpected error happened: ", ex);
            }


            return new OutlookReportEmail();
        }

        [DirectDom("Get Most Recent Mail Item from Folder with Subject")]
        [DirectDomMethod("Get most recent mail item from folder {Folder}, with optional subject {Subject}, only unread {Only Unread}")]
        [MethodDescription("Will return a most recent mail item from given folder. You can supply subject if you wish.")]
        public static OutlookMailItem GetMostRecentMailItem(string folderPath, string subject, bool onlyUnread)
        {

            if (IsOutlookRestarted())
            {
                TryInitializeOutlookApplication();
            }

            try
            {
                if (IsOutlookInstalled)
                {
                    MAPIFolder folderFromFolderPath = GetFolderFromFolderPath(folderPath);
                    if (folderFromFolderPath == null)
                    {
                        logger.Error("Folder not found");
                        return new OutlookMailItem();
                    }
                    Items allItems = folderFromFolderPath.Items;
                    Items outlookItems = onlyUnread ? allItems.Restrict("[Unread]=true") : allItems;
                    if (!string.IsNullOrEmpty(subject))
                    {
                        outlookItems = outlookItems.Restrict(string.Format("@SQL=\"urn:schemas:httpmail:subject\" like '%{0}%'", subject));
                    }

                    outlookItems.Sort("[ReceivedTime]", true);

                    object outlookItem = null;
                    MailItem mailItem = null;

                    for (int i = 1; i <= outlookItems.Count; i++)
                    {
                        mailItem = outlookItems[i] as MailItem;
                        if (mailItem != null)
                        {
                            break;
                        }
                    }

                    return new OutlookMailItem(mailItem);
                }
                else
                {
                    logger.Error(outlookNotInstalledErrorMsg);
                }

            }
            catch (System.Exception ex)
            {
                logger.Error("GetMostRecentMailItem(): unexpected error happened: ", ex);
            }

            return new OutlookMailItem();
        }

        private static List<OutlookReportEmail> GetOutlookReportEmailsFromItems(Items items, int maxCount = -1)
        {
            List<OutlookReportEmail> outlookEmails = new List<OutlookReportEmail>();

            object outlookItem = null;

            for (int i = 1; i <= items.Count; i++)
            {
                outlookItem = items[i];
                try
                {
                    ReportItem reportItem = outlookItem as ReportItem;
                    if (reportItem != null)
                    {
                        OutlookReportEmail email = reportItem.ConvertToOutlookEmail();
                        outlookEmails.Add(email);
                        if (outlookEmails.Count == maxCount)
                        {
                            break;
                        }
                    }
                }
                catch (System.Exception)
                {
                }

            }
            return outlookEmails;
        }

        private static List<MailItem> GetOutlookMailItemsFromItems(Items items, int maxCount = -1)
        {
            List<MailItem> outlookEmails = new List<MailItem>();

            object outlookItem = null;

            for (int i = 1; i <= items.Count; i++)
            {
                outlookItem = items[i];
                try
                {
                    MailItem mailItem = outlookItem as MailItem;
                    if (mailItem != null)
                    {
                        outlookEmails.Add(mailItem);
                        if (outlookEmails.Count == maxCount)
                        {
                            break;
                        }
                    }

                }
                catch (System.Exception)
                {
                }

            }
            return outlookEmails;
        }

        private static MAPIFolder GetFolderFromFolderPath(string folderPath)
        {
            string separator = "\\";
            MAPIFolder returnFolder = null;
            try
            {
                if (folderPath.StartsWith("\\\\"))
                {
                    folderPath = folderPath.Remove(0, 2);
                }

                string[] folderPathArray = folderPath.Split(separator.ToCharArray());

                NameSpace session = Application.Session;
                MAPIFolder targetFolder = session.Folders[folderPathArray[0]];

                if (targetFolder != null)
                {
                    for (int i = 1; i <= folderPathArray.GetUpperBound(0); ++i)
                    {
                        MAPIFolder folder = targetFolder.Folders[folderPathArray[i]];
                        if (folder == null)
                        {
                            return returnFolder;
                        }
                        targetFolder = folder;
                    }
                    returnFolder = targetFolder;
                }
                return returnFolder;
            }
            catch (System.Exception)
            {
                return null;
            }
        }

        private static OutlookReportEmail ConvertToOutlookEmail(this ReportItem reportItem)
        {
            OutlookReportEmail returnEmail = new OutlookReportEmail(reportItem);
            returnEmail.Subject = reportItem.Subject;
            returnEmail.ReceivedDateTime = reportItem.LastModificationTime;
            returnEmail.IsRead = !reportItem.UnRead;
            returnEmail.Preview = reportItem.Body;
            returnEmail.Body = reportItem.Body;
            returnEmail.HasAttachments = reportItem?.Attachments != null && reportItem.Attachments.Count > 0;
            returnEmail.Id = reportItem.EntryID;
            returnEmail.AttachmentsInfo = new DirectCollection<OutlookReportEmailAttachment>();
            if (returnEmail.HasAttachments)
            {
                Attachments attachments = reportItem.Attachments;
                for (int i = 1; i <= attachments.Count; i++)
                {
                    Attachment attachment = attachments[i];
                    if (attachment != null)
                    {
                        OutlookReportEmailAttachment outlookAttachmentInfo = new OutlookReportEmailAttachment(attachment);
                        returnEmail.AttachmentsInfo.Add(outlookAttachmentInfo);
                    }
                }
            }

            return returnEmail;
        }

        private static void MarkEmailsAsRead(List<OutlookReportEmail> listEmails)
        {


            foreach (OutlookReportEmail email in listEmails)
            {
                NameSpace mapi = Application.GetNamespace("MAPI");
                object folderObject = mapi.GetItemFromID(email.Id);
                ReportItem reportItemFromId = folderObject as ReportItem;
                if (reportItemFromId != null)
                {
                    reportItemFromId.UnRead = false;
                    ReleaseObject(reportItemFromId);
                }
                else
                {
                    logger.ErrorFormat("no mail item with id {0} was found", email.Id);
                }

                MailItem mailItemFromId = folderObject as MailItem;
                if (mailItemFromId != null)
                {
                    mailItemFromId.UnRead = false;
                    ReleaseObject(mailItemFromId);
                }
                else
                {
                    logger.ErrorFormat("no mail item with id {0} was found", email.Id);
                }
            }
        }

        private static void ReleaseObject(object objectToRelease)
        {
            if (objectToRelease != null)
            {
                Marshal.ReleaseComObject(objectToRelease);
                objectToRelease = null;
            }
        }

        private static long GetMemoryUsage()
        {
            using (var process = Process.GetCurrentProcess())
            {
                // Returns the amount of physical memory, in bytes, allocated for the associated process.
                return process.WorkingSet64 / 1024 / 1024;
            }
        }

        private static double GetCpuLoad()
        {
            var cpuCounter = new PerformanceCounter("Processor", "% Processor Time", "_Total");

            // The first call will always return 0, so you need to call NextValue twice.
            cpuCounter.NextValue();
            Thread.Sleep(1000); // Wait a second to get a proper reading
            return cpuCounter.NextValue();
        }

        private static bool IsOutlookRestarted()
        {
            var outlookProcesses = Process.GetProcessesByName("OUTLOOK");
            foreach (var process in outlookProcesses)
            {
                try
                {
                    if (process.StartTime > classInitializedTime)
                    {
                        logger.Debug("Outlook was restarted! Class init time: " + classInitializedTime.ToString() + " Outlook process start time: " + process.StartTime.ToString());
                        return true;
                    }
                }
                catch
                {
                }
            }
            return false;
        }

        private static bool TryInitializeOutlookApplication()
        {
            if (Application != null)
            {
                IsOutlookInstalled = false;
                Marshal.ReleaseComObject(Application);
                Application = null;
            }

            int maxRetries = 3;
            int retryDelay = 5000;
            int retryCount = 0;
            bool success = false;

            while (!success && retryCount < maxRetries)
            {
                try
                {
                    logger.Debug("Attempting to get the Outlook instance");
                    Application = new Application();
                    IsOutlookInstalled = true;
                    logger.Debug("Outlook instance obtained successfully");
                    classInitializedTime = DateTime.Now;
                    success = true;
                }
                catch (System.Exception ex)
                {
                    IsOutlookInstalled = false;
                    retryCount++;
                    logger.Debug($"Memory usage: {GetMemoryUsage()}, CPU load: {GetCpuLoad()}");
                    logger.Debug($"Retry {retryCount}/{maxRetries} failed to get Outlook instance. Exception: {ex}");

                    if (retryCount < maxRetries)
                    {
                        logger.Debug($"Waiting for {retryDelay / 1000} seconds before next retry.");
                        Thread.Sleep(retryDelay); // Wait for 5 seconds before retrying
                    }
                }
            }

             return success;
        }
    }
}
