using Direct.Interface;
using Direct.Shared;
using Microsoft.Office.Interop.Outlook;
using System;

namespace PAteam.Library.OutlookExtended
{
    [DirectDom("Outlook Report Email", "Microsoft Outlook Extended")]
    [Bindable]
    public class OutlookReportEmail : DirectComponentBase
    {
        private static IDirectLog logger = DirectLogManager.GetLogger("LibraryObjects");
        protected PropertyHolder<string> _Subject = new PropertyHolder<string>("Subject");
        protected PropertyHolder<DateTime> _ReceivedDateTime = new PropertyHolder<DateTime>("ReceivedDateTime");
        protected PropertyHolder<bool> _IsRead = new PropertyHolder<bool>("IsRead");
        protected PropertyHolder<string> _Preview = new PropertyHolder<string>("Preview");
        protected PropertyHolder<string> _Body = new PropertyHolder<string>("Body");
        protected PropertyHolder<bool> _HasAttachments = new PropertyHolder<bool>("HasAttachments");
        protected PropertyHolder<string> _Id = new PropertyHolder<string>("Id");
        protected PropertyHolder<DirectCollection<OutlookReportEmailAttachment>> _AttachmentsInfo = new PropertyHolder<DirectCollection<OutlookReportEmailAttachment>>("AttachmentsInfo")
        {
            TypedValue = new DirectCollection<OutlookReportEmailAttachment>()
        };
        private ReportItem OriginalReportItem;

        [DirectDom("Subject")]
        public string Subject
        {
            get
            {
                return _Subject.TypedValue;
            }
            set
            {
                _Subject.TypedValue = value;
            }
        }

        [DirectDom("Received DateTime")]
        [DesignTimeInfo("Received DateTime")]
        [ReadOnlyProperty]
        public DateTime ReceivedDateTime
        {
            get
            {
                return _ReceivedDateTime.TypedValue;
            }
            set
            {
                _ReceivedDateTime.TypedValue = value;
            }
        }

        [DirectDom("Is Read")]
        [DesignTimeInfo("Is Read")]
        [ReadOnlyProperty]
        public bool IsRead
        {
            get
            {
                return _IsRead.TypedValue;
            }
            set
            {
                _IsRead.TypedValue = value;
            }
        }

        [DirectDom("Body")]
        public string Body
        {
            get
            {
                return _Body.TypedValue;
            }
            set
            {
                _Body.TypedValue = value;
            }
        }

        [DirectDom("Preview")]
        [ReadOnlyProperty]
        public string Preview
        {
            get
            {
                return _Preview.TypedValue;
            }
            set
            {
                _Preview.TypedValue = value;
            }
        }

        [DirectDom("Has Attachments")]
        [DesignTimeInfo("Has Attachments")]
        [ReadOnlyProperty]
        public bool HasAttachments
        {
            get
            {
                return _HasAttachments.TypedValue;
            }
            set
            {
                _HasAttachments.TypedValue = value;
            }
        }

        [DirectDom("Id")]
        [ReadOnlyProperty]
        public string Id
        {
            get
            {
                return _Id.TypedValue;
            }
            set
            {
                _Id.TypedValue = value;
            }
        }

        [DirectDom("Attachments Info")]
        [DesignTimeInfo("Attachments Info")]
        public DirectCollection<OutlookReportEmailAttachment> AttachmentsInfo
        {
            get
            {
                return _AttachmentsInfo.TypedValue;
            }
            set
            {
                _AttachmentsInfo.TypedValue = value;
            }
        }

        public OutlookReportEmail()
        {
        }

        public OutlookReportEmail(IProject project)
        {
        }

        public OutlookReportEmail(ReportItem reportItem)
        {
            OriginalReportItem = reportItem;
            AttachmentsInfo = new DirectCollection<OutlookReportEmailAttachment>();
        }

        [DirectDom("Add Attachment")]
        [DirectDomMethod("Add Attachment Path {Full Path}")]
        [MethodDescription("Add an attachment to this email's attachments.")]
        public void AddAttachment(string fullPath)
        {
            AttachmentsInfo.Add(new OutlookReportEmailAttachment(fullPath));
            HasAttachments = true;
        }

        [DirectDom("Save as file")]
        [DirectDomMethod("Save report mail as a file to the following location location {File Full Path}")]
        [MethodDescription("Saves report email as a file.")]
        public bool SaveAs(string fullPath)
        {
            try
            {
                if (string.IsNullOrEmpty(fullPath))
                {
                    throw new System.Exception("Empty file path");
                }

                OriginalReportItem.SaveAs(fullPath);
                return true;
            }
            catch (System.Exception ex)
            {
                logger.Error(ex.Message);
                return false;
            }
        }
    }
}
