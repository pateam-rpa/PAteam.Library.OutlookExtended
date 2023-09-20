using Direct.Interface;
using Direct.Shared;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Win32;
using System.IO;
using System.Web;

namespace PAteam.Library.OutlookExtended
{
    [DirectDom("Outlook Report Attachment", "Microsoft Outlook Extended")]
    [Bindable]
    public class OutlookReportEmailAttachment : DirectComponentBase
    {
        private static IDirectLog logger = DirectLogManager.GetLogger("LibraryObjects");

        protected PropertyHolder<string> _Id = new PropertyHolder<string>("Id");

        protected PropertyHolder<string> _Name = new PropertyHolder<string>("Name");

        protected PropertyHolder<bool> _IsInline = new PropertyHolder<bool>("IsInline");

        protected PropertyHolder<int> _Size = new PropertyHolder<int>("Size");

        protected PropertyHolder<string> _ContentType = new PropertyHolder<string>("ContentType");

        public readonly string ATTACHMENT_FULL_PATH;

        private readonly Attachment OriginalAttachment;

        public byte[] ContentBytes { get; private set; }

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

        [DirectDom("Name")]
        public string Name
        {
            get
            {
                return _Name.TypedValue;
            }
            set
            {
                _Name.TypedValue = value;
            }
        }

        [DirectDom("Is Inline")]
        [DesignTimeInfo("Is Inline")]
        [ReadOnlyProperty]
        public bool IsInline
        {
            get
            {
                return _IsInline.TypedValue;
            }
            set
            {
                _IsInline.TypedValue = value;
            }
        }

        [DirectDom("Size")]
        [ReadOnlyProperty]
        public int Size
        {
            get
            {
                return _Size.TypedValue;
            }
            set
            {
                _Size.TypedValue = value;
            }
        }

        [DirectDom("ContentType")]
        [DesignTimeInfo("ContentType")]
        public string ContentType
        {
            get
            {
                return _ContentType.TypedValue;
            }
            set
            {
                _ContentType.TypedValue = value;
            }
        }

        [DirectDom("Save as file")]
        [DirectDomMethod("Save attachment as a file to the following loaction: {Full File Path}")]
        [MethodDescription("Saves attachment as file")]
        public bool SaveAsFile(string filePath)
        {
            try
            {
                if (string.IsNullOrEmpty(filePath)) 
                {
                    throw new System.Exception("Empty file path");
                }

                OriginalAttachment.SaveAsFile(filePath);
                return true;
            }
            catch (System.Exception ex)
            {
                logger.Error(ex.Message);
                return false;
            }
        }

        public OutlookReportEmailAttachment()
        {
        }

        public OutlookReportEmailAttachment(IProject project)
        {
        }

        public OutlookReportEmailAttachment(Attachment attachment)
        {
            OriginalAttachment = attachment;
            Id = attachment.Index.ToString();
            Name = attachment.FileName;
            IsInline = false;
            Size = attachment.Size;
            ContentType = GetMimeType(attachment.FileName);
        }

        public OutlookReportEmailAttachment(string id, string name, bool isInline, int size, string contentType)
        {
            Id = id;
            Name = name;
            IsInline = isInline;
            Size = size;
            ContentType = contentType;
        }

        public OutlookReportEmailAttachment(string fileFullPath)
        {
            if (!File.Exists(fileFullPath))
            {
                return;
            }

            IsInline = false;
            Size = 0;
            ContentType = "application/unknown";
            Name = Path.GetFileName(fileFullPath);
            ATTACHMENT_FULL_PATH = fileFullPath;
            ContentType = GetMimeType(fileFullPath);
            try
            {
                Size = (int)new FileInfo(fileFullPath).Length;
            }
            catch (System.Exception ex)
            {
                logger.Error(ex.Message);
            }

            using (FileStream input = File.OpenRead(fileFullPath))
            {
                using (BinaryReader binaryReader = new BinaryReader(input))
                    ContentBytes = binaryReader.ReadBytes((int)input.Length);
            }
        }

        private static string GetMimeType(string fileName)
        {
            return MimeMapping.GetMimeMapping(fileName);

            //RegistryKey registryKey = Registry.ClassesRoot.OpenSubKey(name);
            //if (registryKey != null && registryKey.GetValue("Content Type") != null)
            //{
            //    result = registryKey.GetValue("Content Type").ToString();
            //}

            //return result;
        }
    }
}

