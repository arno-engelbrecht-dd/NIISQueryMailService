using log4net;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MailKit.Security;
using MimeKit;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace NIISQueryMailService
{
    public partial class NIISQueryMailService : ServiceBase
    {
        private static readonly ILog _log = LogManager.GetLogger(typeof(NIISQueryMailService));

        Thread _mainThread = null;
        public static bool _bAbortingThread = false;
        static EventLog _processEventLog = null;

        private static string mailServer;
        private static int mailPort;
        private static string mailUser;
        private static string mailAddress;
        private static string mailPassword;

        private static string documentsDirectory;

        private static string MailServer
        {
            get
            {
                if (!string.IsNullOrEmpty(mailServer))
                    return mailServer;
                try
                {
                    var appSettings = new System.Configuration.AppSettingsReader();
                    mailServer = (string)(appSettings.GetValue("MailServer", typeof(string)));
                    mailPort = (int)(appSettings.GetValue("MailPort", typeof(int)));
                    mailUser = (string)(appSettings.GetValue("MailUser", typeof(string)));
                    mailAddress = (string)(appSettings.GetValue("MailAddress", typeof(string)));
                    mailPassword = (string)(appSettings.GetValue("MailPassword", typeof(string)));
                }
                catch (Exception ex)
                {
                    _log.Error(ex);
                }
                return mailServer;
            }
        }
        private static int MailPort
        {
            get
            {
                if (mailPort > 0)
                    return mailPort;
                try
                {
                    var appSettings = new System.Configuration.AppSettingsReader();
                    mailServer = (string)(appSettings.GetValue("MailServer", typeof(string)));
                    mailPort = (int)(appSettings.GetValue("MailPort", typeof(int)));
                    mailUser = (string)(appSettings.GetValue("MailUser", typeof(string)));
                    mailAddress = (string)(appSettings.GetValue("MailAddress", typeof(string)));
                    mailPassword = (string)(appSettings.GetValue("MailPassword", typeof(string)));
                }
                catch (Exception ex)
                {
                    _log.Error(ex);
                }
                return mailPort;
            }
        }
        private static string MailUser
        {
            get
            {
                if (!string.IsNullOrEmpty(mailUser))
                    return mailUser;
                try
                {
                    var appSettings = new System.Configuration.AppSettingsReader();
                    mailServer = (string)(appSettings.GetValue("MailServer", typeof(string)));
                    mailPort = (int)(appSettings.GetValue("MailPort", typeof(int)));
                    mailUser = (string)(appSettings.GetValue("MailUser", typeof(string)));
                    mailAddress = (string)(appSettings.GetValue("MailAddress", typeof(string)));
                    mailPassword = (string)(appSettings.GetValue("MailPassword", typeof(string)));
                }
                catch (Exception ex)
                {
                    _log.Error(ex);
                }
                return mailUser;
            }
        }
        private static string MailAddress
        {
            get
            {
                if (!string.IsNullOrEmpty(mailAddress))
                    return mailAddress;
                try
                {
                    var appSettings = new System.Configuration.AppSettingsReader();
                    mailServer = (string)(appSettings.GetValue("MailServer", typeof(string)));
                    mailPort = (int)(appSettings.GetValue("MailPort", typeof(int)));
                    mailUser = (string)(appSettings.GetValue("MailUser", typeof(string)));
                    mailAddress = (string)(appSettings.GetValue("MailAddress", typeof(string)));
                    mailPassword = (string)(appSettings.GetValue("MailPassword", typeof(string)));
                }
                catch (Exception ex)
                {
                    _log.Error(ex);
                }
                return mailAddress;
            }
        }
        private static string MailPassword
        {
            get
            {
                if (!string.IsNullOrEmpty(mailPassword))
                    return mailPassword;
                try
                {
                    var appSettings = new System.Configuration.AppSettingsReader();
                    mailServer = (string)(appSettings.GetValue("MailServer", typeof(string)));
                    mailPort = (int)(appSettings.GetValue("MailPort", typeof(int)));
                    mailUser = (string)(appSettings.GetValue("MailUser", typeof(string)));
                    mailAddress = (string)(appSettings.GetValue("MailAddress", typeof(string)));
                    mailPassword = (string)(appSettings.GetValue("MailPassword", typeof(string)));
                }
                catch (Exception ex)
                {
                    _log.Error(ex);
                }
                return mailPassword;
            }
        }

        private static string DocumentsDirectory
        {
            get
            {
                if (!string.IsNullOrEmpty(documentsDirectory))
                    return documentsDirectory;
                try
                {
                    var appSettings = new System.Configuration.AppSettingsReader();
                    documentsDirectory = (string)(appSettings.GetValue("DocumentsDirectory", typeof(string)));
                }
                catch (Exception ex)
                {
                    _log.Error(ex);
                }
                return documentsDirectory;
            }
        }

        public NIISQueryMailService()
        {
            InitializeComponent();

            this.ServiceName = "NIIS Query Mail service";

            if (!EventLog.SourceExists("NIISQueryMailService"))
            {
                EventLog.CreateEventSource("NIISQueryMailService", "NIISQueryMailService");
            }
            this.EventLog.Source = "NIISQueryMailService";
            this.EventLog.Log = "NIISQueryMailService";

            _processEventLog = this.EventLog;

            this.CanHandlePowerEvent = true;
            this.CanHandleSessionChangeEvent = true;
            this.CanPauseAndContinue = true;
            this.CanShutdown = true;
            this.CanStop = true;

            log4net.Config.XmlConfigurator.Configure();
        }

        public static void WriteLog(string message)
        {
            _processEventLog.WriteEntry(message);
        }

        protected override void OnStart(string[] args)
        {
            WriteLog("Starting NIIS Query Mail service");
            _log.Info("Starting NIIS Query Mail service");

            try
            {
                _mainThread = new Thread(new ThreadStart(DoWork));
                _mainThread.Start();
            }
            catch (Exception ex)
            {
                _log.Error(ex);
                WriteLog("Error starting service: " + ex.Message);
                this.Stop();
                throw ex;
            }
        }

        protected override void OnStop()
        {
            WriteLog("Stopping NIIS Query Mail service");
            _log.Info("Stopping NIIS Query Mail service");

            _bAbortingThread = true;

            base.OnStop();

            if (_mainThread != null)
            {
                _mainThread.Join();
            }
        }

        public static void DoWork()
        {
            _bAbortingThread = false;

            while (!_bAbortingThread)
            {
                try
                {
                    GetMails(MailServer, MailPort, MailUser, MailPassword);
                
                    // Sleep for 5 seconds
                    Thread.Sleep(5000);
                }
                catch (Exception ex)
                {
                    _log.Error(ex);
                    _log.Error(ex.StackTrace);
                    WriteLog("Error in DoWork: " + ex.Message);
                }
            }
        }

        private static void GetMails(string server, int port, string user, string password)
        {
            using (var client = new ImapClient())
            {
                client.Connect(server, port, SecureSocketOptions.None);
                client.Authenticate(user, password);

                client.Inbox.Open(FolderAccess.ReadWrite);

                var uids = client.Inbox.Search(SearchQuery.NotSeen);

                var db = new SqlDB();
                try
                {
                    foreach (var uid in uids)
                    {
                        if (_bAbortingThread)
                        {
                            client.Disconnect(true);
                            return;
                        }

                        var msg = client.Inbox.GetMessage(uid);

                        var msgId = msg.MessageId;
                        var fromAddress = msg.From.OfType<MailboxAddress>().Single().Address;

                        client.Inbox.SetFlags(uid, MessageFlags.Seen, true);

                        if (string.IsNullOrEmpty(msgId))
                            continue;

                        var strCount = db.ReadSingleValue("SELECT COUNT(*) FROM QueryMails WHERE MailID=@1", new object[] { msgId }).ToString();
                        if (!strCount.Equals("0"))
                        {
                            continue;
                        }

                        string subject = msg.Subject;
                        var txt = msg.TextBody;
                        if (string.IsNullOrEmpty(txt))
                        {
                            txt = msg.HtmlBody;
                        }
                        if (string.IsNullOrEmpty(txt))
                        {
                            txt = string.Empty;
                        }

                        db.ExecuteSQL_NoClose("INSERT INTO QueryMails (MailID, CreateDate, FromMail, Status, Email, Subject, MailContents) VALUES (@1,@2,@3,@4,@3,@5,@6)",
                                                new object[] { msgId, DateTime.Now, fromAddress, "Pending", subject, txt });

                        var attachments = msg.BodyParts.OfType<MimePart>().Where(part => !string.IsNullOrEmpty(part.FileName)).ToList();
                        if ((attachments == null) || (attachments.Count == 0))
                        {
                            attachments = new List<MimePart>();
                            var atts = msg.Attachments.ToList();
                            if ((atts != null) && (atts.Count > 0))
                            {
                                foreach (MimeEntity att in atts)
                                {
                                    if (att.IsAttachment)
                                    {
                                        var mp = ((MimePart)att);
                                        attachments.Add(mp);
                                    }
                                }
                            }
                        }

                        if ((attachments != null) && (attachments.Count > 0))
                        {
                            attachments.RemoveAll(a => !a.FileName.ToLower().EndsWith("pdf"));

                            attachments.RemoveAll(a => a.ContentDisposition.Size < 100);
                            attachments.RemoveAll(a => !a.IsAttachment);
                        }

                        if (attachments.Count == 0)
                        {
                            continue;
                        }

                        var mailID = db.ReadSingleValue("SELECT ID FROM QueryMails WHERE MailID=@1", new object[] { msgId }).ToString();

                        var subPath = DocumentsDirectory;
                        if (!Directory.Exists(subPath))
                            Directory.CreateDirectory(subPath);

                        foreach (MimeEntity att in attachments)
                        {
                            if (att.IsAttachment)
                            {
                                var mp = ((MimePart)att);
                                var fileName = Path.GetFileName(mp.FileName);
                                var completedFileName = DateTime.Now.ToFileTime() + "_" + fileName;
                                using (var stream = File.Create(Path.Combine(subPath, completedFileName)))
                                {
                                    mp.Content.DecodeTo(stream);
                                }
                                db.ExecuteSQL_NoClose("INSERT INTO QueryMailAttachments(QueryID, FileName, Name, Description, IsDeleted) VALUES (@1,@2,@3,@4,@5)",
                                                        new object[] { mailID, completedFileName, Path.GetFileName(fileName), string.Empty, false });
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    _log.Error(ex);
                }
                finally
                {
                    db.CloseConnection();
                }

                client.Disconnect(true);

                return;
            }
        }
    }
}
