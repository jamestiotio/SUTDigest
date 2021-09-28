using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using Outlook = Microsoft.Office.Interop.Outlook;

[assembly: DefaultDllImportSearchPaths(DllImportSearchPath.System32)]

namespace SUTDigest
{
    public class OutlookWrapper
    {
        private Outlook.Application application;
        private Outlook.NameSpace nameSpace;
        private Outlook.Account account;
        private Outlook.MAPIFolder folder;
        private Outlook.MailItem mail;
        private Outlook.OlDefaultFolders defaultFolderSaved;
        private bool isOutlookVisible = false;

        private const int SW_MAXIMIZE = 3;
        private const int SW_MINIMIZE = 6;

        // This path is for Microsoft Outlook v16.0
        private const string MS_OUTLOOK_EXECUTABLE_FILE_PATH = "C:\\Program Files\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE";

        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        private static void ReleaseComObject(object obj)
        {
            if (obj != null)
            {
                // Take note that this method is only supported on Windows.
#pragma warning disable CA1416 // Validate platform compatibility
                Marshal.ReleaseComObject(obj);
#pragma warning restore CA1416 // Validate platform compatibility
                obj = null;
            }
        }

        public void LaunchOutlook()
        {
            try
            {
                // Check whether there is an Outlook process running.
                if (Process.GetProcessesByName("OUTLOOK").Length == 0)
                {
                    // If not, create a new instance of Outlook and log on to the default profile.
                    var outlookApp = new Process
                    {
                        // Specify the Microsoft Outlook executable file location here
                        StartInfo = new ProcessStartInfo(MS_OUTLOOK_EXECUTABLE_FILE_PATH)
                    };
                    // outlookApp.StartInfo.Verb = "runas";    // Only required if Outlook is to be launched with Administrator rights (preferably not for normal regular instances due to principle of least privilege).
                    outlookApp.Start();
                    // Following steps are needed to ensure that the Outlook Application instance is added
                    // to Running Object Table (ROT), so that "AlternativeMarshal.GetActiveObject()" won't fail.
                    Thread.Sleep(20000);
                    outlookApp.WaitForInputIdle();
                    ShowWindow(outlookApp.MainWindowHandle, SW_MINIMIZE);       // Minimize Window to force addition to ROT
                    Thread.Sleep(3000);
                    ShowWindow(outlookApp.MainWindowHandle, SW_MAXIMIZE);       // Restore Window
                    Thread.Sleep(2000);

                    isOutlookVisible = true;
                }
            }
            catch (Exception ex)
            {
                ReleaseComObject(application);
                ReleaseComObject(nameSpace);
                ReleaseComObject(account);
                ReleaseComObject(folder);
                ReleaseComObject(mail);
            }
        }

        public void SelectProfile()
        {
            try
            {
                // Check whether there is an Outlook process running.
                if (Process.GetProcessesByName("OUTLOOK").Any())
                {
                    // If so, use the GetActiveObject method to obtain the process and cast it to an Application object.
                    // We P/Invoke an alternative custom function since the specific default GetActiveObject() method API is gone in .NET 5/.NET Core.
                    // For more information: https://docs.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-getactiveobject
                    application = AlternativeMarshal.GetActiveObject("Outlook.Application") as Outlook.Application;
                    nameSpace = application.GetNamespace("MAPI");
                    isOutlookVisible = true;
                }
                else
                {
                    // If not, create a new instance of Outlook and log on to the provided profile.
                    // This is usually the recommended method since Outlook should be a singleton (only one instance of outlook.exe running at all times).
                    // Creating a new object should return the existing object if Outlook is already running.
                    application = new Outlook.Application();
                    nameSpace = application.GetNamespace("MAPI");
                    isOutlookVisible = false;
                }

                // This ensures that Outlook is fully initialized.
                // If Outlook is already running, the call will do nothing.
                nameSpace.Logon(string.Empty, string.Empty, false, true);
            }
            catch (Exception ex)
            {
                ReleaseComObject(application);
                ReleaseComObject(nameSpace);
                ReleaseComObject(account);
                ReleaseComObject(folder);
                ReleaseComObject(mail);
            }
        }

        public static Outlook.Folder GetFolder(string folderPath, Outlook.Application outlookApp)
        {
            Outlook.Folder folder;
            string backslash = @"\";
            try
            {
                if (folderPath.StartsWith(@"\\"))
                {
                    folderPath = folderPath.Remove(0, 2);
                }

                String[] folders =
                    folderPath.Split(backslash.ToCharArray());
                folder =
                    outlookApp.Session.Folders[folders[0]]
                    as Outlook.Folder;

                if (folder != null)
                {
                    for (int i = 1; i <= folders.GetUpperBound(0); i++)
                    {
                        Outlook.Folders subFolders = folder.Folders;
                        folder = subFolders[folders[i]]
                            as Outlook.Folder;

                        if (folder == null)
                        {
                            return null;
                        }
                    }
                }
                return folder;
            }
            catch
            {
                return null;
            }
        }

        private bool RuleExist(string ruleName, Outlook.Rules rules)
        {
            bool exist = false;

            for (int i = 1; i <= rules.Count; i++)
            {
                Outlook.Rule rule = rules[i];
                if (rule.Name == ruleName)
                {
                    ReleaseComObject(rule);
                    exist = true;
                    break;
                }
                ReleaseComObject(rule);
            }
            return exist;
        }

        public void CreateFolders()
        {
            Outlook.NameSpace session = null;
            Outlook.Store store = null;
            Outlook.MAPIFolder rootFolder = null;
            Outlook.Folders rootFolderFolders = null;
            Outlook.MAPIFolder othersFolder = null;
            Outlook.Folders othersSubfolders = null;
            Outlook.MAPIFolder studentClubsFolder = null;
            Outlook.MAPIFolder marketingFolder = null;
            Outlook.MAPIFolder cdcFolder = null;
            Outlook.MAPIFolder whatzNewzFolder = null;
            Outlook.MAPIFolder libraryFolder = null;
            Outlook.MAPIFolder studentGovFolder = null;
            Outlook.MAPIFolder allStudentsFolder = null;
            Outlook.MAPIFolder newsCoverageFolder = null;
            Outlook.MAPIFolder socialMediaReportFolder = null;
            Outlook.MAPIFolder researchFolder = null;

            try
            {
                session = application.Session;
                store = session.DefaultStore;
                rootFolder = store.GetRootFolder();

                // Folders are identified by name
                othersFolder = GetFolder(rootFolder.FolderPath + @"\Others", application);

                if (othersFolder == null)
                {
                    rootFolderFolders = rootFolder.Folders;
                    othersFolder = rootFolderFolders.Add("Others", Outlook.OlDefaultFolders.olFolderInbox);
                }

                othersSubfolders = othersFolder.Folders;

                studentClubsFolder = GetFolder(rootFolder.FolderPath + @"\Others\Student Clubs", application);

                if (studentClubsFolder == null)
                {
                    studentClubsFolder = othersSubfolders.Add("Student Clubs", Outlook.OlDefaultFolders.olFolderInbox);
                }

                marketingFolder = GetFolder(rootFolder.FolderPath + @"\Others\Office of Marketing", application);

                if (marketingFolder == null)
                {
                    marketingFolder = othersSubfolders.Add("Office of Marketing", Outlook.OlDefaultFolders.olFolderInbox);
                }

                cdcFolder = GetFolder(rootFolder.FolderPath + @"\Others\Career Development Centre", application);

                if (cdcFolder == null)
                {
                    cdcFolder = othersSubfolders.Add("Career Development Centre", Outlook.OlDefaultFolders.olFolderInbox);
                }

                whatzNewzFolder = GetFolder(rootFolder.FolderPath + @"\Others\Whatz Newz", application);

                if (whatzNewzFolder == null)
                {
                    whatzNewzFolder = othersSubfolders.Add("Whatz Newz", Outlook.OlDefaultFolders.olFolderInbox);
                }

                libraryFolder = GetFolder(rootFolder.FolderPath + @"\Others\Library", application);

                if (libraryFolder == null)
                {
                    libraryFolder = othersSubfolders.Add("Library", Outlook.OlDefaultFolders.olFolderInbox);
                }

                studentGovFolder = GetFolder(rootFolder.FolderPath + @"\Others\Student Government", application);

                if (studentGovFolder == null)
                {
                    studentGovFolder = othersSubfolders.Add("Student Government", Outlook.OlDefaultFolders.olFolderInbox);
                }

                allStudentsFolder = GetFolder(rootFolder.FolderPath + @"\Others\All Students", application);

                if (allStudentsFolder == null)
                {
                    allStudentsFolder = othersSubfolders.Add("All Students", Outlook.OlDefaultFolders.olFolderInbox);
                }

                newsCoverageFolder = GetFolder(rootFolder.FolderPath + @"\Others\News Coverage", application);

                if (newsCoverageFolder == null)
                {
                    newsCoverageFolder = othersSubfolders.Add("News Coverage", Outlook.OlDefaultFolders.olFolderInbox);
                }

                socialMediaReportFolder = GetFolder(rootFolder.FolderPath + @"\Others\Social Media Report", application);

                if (socialMediaReportFolder == null)
                {
                    socialMediaReportFolder = othersSubfolders.Add("Social Media Report", Outlook.OlDefaultFolders.olFolderInbox);
                }

                researchFolder = GetFolder(rootFolder.FolderPath + @"\Others\Office of Research", application);

                if (researchFolder == null)
                {
                    researchFolder = othersSubfolders.Add("Office of Research", Outlook.OlDefaultFolders.olFolderInbox);
                }
            }
            catch (Exception ex)
            {
                // If the user is not connected to Microsoft Exchange or if they are disconnected, an exception will be raised.
                Debug.Write(ex.Message);
            }
            finally
            {
                if (rootFolderFolders != null)
                    ReleaseComObject(rootFolderFolders);
                if (rootFolder != null)
                    ReleaseComObject(rootFolder);
                if (othersFolder != null)
                    ReleaseComObject(othersFolder);
                if (othersSubfolders != null)
                    ReleaseComObject(othersSubfolders);
                if (studentClubsFolder != null)
                    ReleaseComObject(studentClubsFolder);
                if (marketingFolder != null)
                    ReleaseComObject(marketingFolder);
                if (cdcFolder != null)
                    ReleaseComObject(cdcFolder);
                if (whatzNewzFolder != null)
                    ReleaseComObject(whatzNewzFolder);
                if (libraryFolder != null)
                    ReleaseComObject(libraryFolder);
                if (studentGovFolder != null)
                    ReleaseComObject(studentGovFolder);
                if (allStudentsFolder != null)
                    ReleaseComObject(allStudentsFolder);
                if (newsCoverageFolder != null)
                    ReleaseComObject(newsCoverageFolder);
                if (socialMediaReportFolder != null)
                    ReleaseComObject(socialMediaReportFolder);
                if (researchFolder != null)
                    ReleaseComObject(researchFolder);
                if (store != null)
                    ReleaseComObject(store);
                if (session != null)
                    ReleaseComObject(session);
            }
        }

        public void CreateRules()
        {
            Outlook.NameSpace session = null;
            Outlook.Store store = null;
            Outlook.Rules rules = null;
            Outlook.MAPIFolder rootFolder = null;
            Outlook.Folders rootFolderFolders = null;

            try
            {
                session = application.Session;
                store = session.DefaultStore;
                rules = store.GetRules();
                rootFolder = store.GetRootFolder();

                // Rules are identified by name
                string studentClubRuleName = "Emails from Student Clubs";

                if (!RuleExist(studentClubRuleName, rules))
                {
                    Outlook.MAPIFolder destinationFolder = GetFolder(rootFolder.FolderPath + @"\Others\Student Clubs", application);

                    Outlook.Rule rule = rules.Create(studentClubRuleName, Outlook.OlRuleType.olRuleReceive);
                    Outlook.RuleConditions ruleConditions = rule.Conditions;

                    Outlook.AddressRuleCondition senderAddressRuleCondition = ruleConditions.SenderAddress;
                    senderAddressRuleCondition.Address = new string[] { "club.sutd.edu.sg" };
                    senderAddressRuleCondition.Enabled = true;

                    Outlook.RuleActions ruleActions = rule.Actions;
                    Outlook.MoveOrCopyRuleAction moveRuleAction = ruleActions.MoveToFolder;
                    moveRuleAction.Folder = destinationFolder;
                    moveRuleAction.Enabled = true;

                    // This ensures that each rule is independent of each other
                    ruleActions.Stop.Enabled = true;

                    rules.Save(true);
                }

                string marketingRuleName = "Emails from Office of Marketing";

                if (!RuleExist(marketingRuleName, rules))
                {
                    Outlook.MAPIFolder destinationFolder = GetFolder(rootFolder.FolderPath + @"\Others\Office of Marketing", application);

                    Outlook.Rule rule = rules.Create(marketingRuleName, Outlook.OlRuleType.olRuleReceive);
                    Outlook.RuleConditions ruleConditions = rule.Conditions;

                    Outlook.AddressRuleCondition senderAddressRuleCondition = ruleConditions.SenderAddress;
                    senderAddressRuleCondition.Address = new string[] { "marcom.sutd.edu.sg" };
                    senderAddressRuleCondition.Enabled = true;

                    Outlook.RuleActions ruleActions = rule.Actions;
                    Outlook.MoveOrCopyRuleAction moveRuleAction = ruleActions.MoveToFolder;
                    moveRuleAction.Folder = destinationFolder;
                    moveRuleAction.Enabled = true;

                    ruleActions.Stop.Enabled = true;

                    rules.Save(true);
                }

                string cdcRuleName = "Emails from Career Development Centre";

                if (!RuleExist(cdcRuleName, rules))
                {
                    Outlook.MAPIFolder destinationFolder = GetFolder(rootFolder.FolderPath + @"\Others\Career Development Centre", application);

                    Outlook.Rule rule = rules.Create(cdcRuleName, Outlook.OlRuleType.olRuleReceive);
                    Outlook.RuleConditions ruleConditions = rule.Conditions;

                    Outlook.ToOrFromRuleCondition senderAddressRuleCondition = ruleConditions.From;
                    senderAddressRuleCondition.Recipients.Add("careers@sutd.edu.sg");
                    senderAddressRuleCondition.Recipients.ResolveAll();
                    senderAddressRuleCondition.Enabled = true;

                    Outlook.RuleActions ruleActions = rule.Actions;
                    Outlook.MoveOrCopyRuleAction moveRuleAction = ruleActions.MoveToFolder;
                    moveRuleAction.Folder = destinationFolder;
                    moveRuleAction.Enabled = true;

                    ruleActions.Stop.Enabled = true;

                    rules.Save(true);
                }

                string whatzNewzRuleName = "Whatz Newz Emails";

                if (!RuleExist(whatzNewzRuleName, rules))
                {
                    Outlook.MAPIFolder destinationFolder = GetFolder(rootFolder.FolderPath + @"\Others\Whatz Newz", application);

                    Outlook.Rule rule = rules.Create(whatzNewzRuleName, Outlook.OlRuleType.olRuleReceive);
                    Outlook.RuleConditions ruleConditions = rule.Conditions;

                    Outlook.TextRuleCondition subjectTextRuleCondition = ruleConditions.Subject;
                    subjectTextRuleCondition.Text = new string[] { "Whatz Newz by Office of Student Life" };
                    subjectTextRuleCondition.Enabled = true;

                    Outlook.RuleActions ruleActions = rule.Actions;
                    Outlook.MoveOrCopyRuleAction moveRuleAction = ruleActions.MoveToFolder;
                    moveRuleAction.Folder = destinationFolder;
                    moveRuleAction.Enabled = true;

                    ruleActions.Stop.Enabled = true;

                    rules.Save(true);
                }

                string libraryRuleName = "Emails from Library";

                if (!RuleExist(libraryRuleName, rules))
                {
                    Outlook.MAPIFolder destinationFolder = GetFolder(rootFolder.FolderPath + @"\Others\Library", application);

                    Outlook.Rule rule = rules.Create(libraryRuleName, Outlook.OlRuleType.olRuleReceive);
                    Outlook.RuleConditions ruleConditions = rule.Conditions;

                    Outlook.ToOrFromRuleCondition senderAddressRuleCondition = ruleConditions.From;
                    senderAddressRuleCondition.Recipients.Add("library@sutd.edu.sg");
                    senderAddressRuleCondition.Recipients.ResolveAll();
                    senderAddressRuleCondition.Enabled = true;

                    Outlook.RuleActions ruleActions = rule.Actions;
                    Outlook.MoveOrCopyRuleAction moveRuleAction = ruleActions.MoveToFolder;
                    moveRuleAction.Folder = destinationFolder;
                    moveRuleAction.Enabled = true;

                    ruleActions.Stop.Enabled = true;

                    rules.Save(true);
                }

                string studentGovRuleName = "Emails from Student Government";

                if (!RuleExist(studentGovRuleName, rules))
                {
                    Outlook.MAPIFolder destinationFolder = GetFolder(rootFolder.FolderPath + @"\Others\Student Government", application);

                    Outlook.Rule rule = rules.Create(studentGovRuleName, Outlook.OlRuleType.olRuleReceive);
                    Outlook.RuleConditions ruleConditions = rule.Conditions;

                    Outlook.AddressRuleCondition senderAddressRuleCondition = ruleConditions.SenderAddress;
                    senderAddressRuleCondition.Address = new string[] { "studentgov.sutd.edu.sg" };
                    senderAddressRuleCondition.Enabled = true;

                    Outlook.RuleActions ruleActions = rule.Actions;
                    Outlook.MoveOrCopyRuleAction moveRuleAction = ruleActions.MoveToFolder;
                    moveRuleAction.Folder = destinationFolder;
                    moveRuleAction.Enabled = true;

                    ruleActions.Stop.Enabled = true;

                    rules.Save(true);
                }

                string newsCoverageRuleName = "News Coverage Emails";

                if (!RuleExist(newsCoverageRuleName, rules))
                {
                    Outlook.MAPIFolder destinationFolder = GetFolder(rootFolder.FolderPath + @"\Others\News Coverage", application);

                    Outlook.Rule rule = rules.Create(newsCoverageRuleName, Outlook.OlRuleType.olRuleReceive);
                    Outlook.RuleConditions ruleConditions = rule.Conditions;

                    Outlook.TextRuleCondition subjectTextRuleCondition = ruleConditions.Subject;
                    subjectTextRuleCondition.Text = new string[] { "Summary of coverage: Highlights of SUTD" };
                    subjectTextRuleCondition.Enabled = true;

                    Outlook.RuleActions ruleActions = rule.Actions;
                    Outlook.MoveOrCopyRuleAction moveRuleAction = ruleActions.MoveToFolder;
                    moveRuleAction.Folder = destinationFolder;
                    moveRuleAction.Enabled = true;

                    ruleActions.Stop.Enabled = true;

                    rules.Save(true);
                }

                // This rule is client-only
                string allStudentsRuleName = "Emails Sent To All Students Without BCC";

                if (!RuleExist(allStudentsRuleName, rules))
                {
                    Outlook.MAPIFolder destinationFolder = GetFolder(rootFolder.FolderPath + @"\Others\All Students", application);

                    Outlook.Rule rule = rules.Create(allStudentsRuleName, Outlook.OlRuleType.olRuleReceive);
                    Outlook.RuleConditions ruleConditions = rule.Conditions;

                    Outlook.ToOrFromRuleCondition senderAddressRuleCondition = ruleConditions.SentTo;
                    senderAddressRuleCondition.Recipients.Add("*AllStudents@sutd.edu.sg");
                    senderAddressRuleCondition.Recipients.ResolveAll();
                    senderAddressRuleCondition.Enabled = true;

                    Outlook.RuleActions ruleActions = rule.Actions;
                    Outlook.MoveOrCopyRuleAction moveRuleAction = ruleActions.MoveToFolder;
                    moveRuleAction.Folder = destinationFolder;
                    moveRuleAction.Enabled = true;

                    ruleActions.Stop.Enabled = true;

                    rules.Save(true);
                }

                string socialMediaReportRuleName = "Social Media Report Emails";

                if (!RuleExist(socialMediaReportRuleName, rules))
                {
                    Outlook.MAPIFolder destinationFolder = GetFolder(rootFolder.FolderPath + @"\Others\Social Media Report", application);

                    Outlook.Rule rule = rules.Create(socialMediaReportRuleName, Outlook.OlRuleType.olRuleReceive);
                    Outlook.RuleConditions ruleConditions = rule.Conditions;

                    Outlook.TextRuleCondition subjectTextRuleCondition = ruleConditions.Subject;
                    subjectTextRuleCondition.Text = new string[] { "SUTD's Social Media Report" };
                    subjectTextRuleCondition.Enabled = true;

                    Outlook.RuleActions ruleActions = rule.Actions;
                    Outlook.MoveOrCopyRuleAction moveRuleAction = ruleActions.MoveToFolder;
                    moveRuleAction.Folder = destinationFolder;
                    moveRuleAction.Enabled = true;

                    ruleActions.Stop.Enabled = true;

                    rules.Save(true);
                }

                string researchRuleName = "Emails from Office of Research";

                if (!RuleExist(researchRuleName, rules))
                {
                    Outlook.MAPIFolder destinationFolder = GetFolder(rootFolder.FolderPath + @"\Others\Office of Research", application);

                    Outlook.Rule rule = rules.Create(researchRuleName, Outlook.OlRuleType.olRuleReceive);
                    Outlook.RuleConditions ruleConditions = rule.Conditions;

                    Outlook.ToOrFromRuleCondition senderAddressRuleCondition = ruleConditions.From;
                    senderAddressRuleCondition.Recipients.Add("research@sutd.edu.sg");
                    senderAddressRuleCondition.Recipients.ResolveAll();
                    senderAddressRuleCondition.Enabled = true;

                    Outlook.RuleActions ruleActions = rule.Actions;
                    Outlook.MoveOrCopyRuleAction moveRuleAction = ruleActions.MoveToFolder;
                    moveRuleAction.Folder = destinationFolder;
                    moveRuleAction.Enabled = true;

                    ruleActions.Stop.Enabled = true;

                    rules.Save(true);
                }
            }
            catch (Exception ex)
            {
                // If the user is not connected to Microsoft Exchange or if they are disconnected, an exception will be raised.
                Debug.Write(ex.Message);
            }
            finally
            {
                if (rootFolderFolders != null)
                    ReleaseComObject(rootFolderFolders);
                if (rootFolder != null)
                    ReleaseComObject(rootFolder);
                if (rules != null)
                    ReleaseComObject(rules);
                if (store != null)
                    ReleaseComObject(store);
                if (session != null)
                    ReleaseComObject(session);
            }
        }
    }
}
