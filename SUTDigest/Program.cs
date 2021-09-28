using System;
using System.Media;

namespace SUTDigest
{
    class Program
    {
        static void Main(string[] args)
        {
            OutlookWrapper outlookWrapper = new();
            // Check if there is an Outlook process to hook to and launch a new one if not
            outlookWrapper.LaunchOutlook();
            // Login with default profile or existing session
            outlookWrapper.SelectProfile();
            // Create folders
            outlookWrapper.CreateFolders();
            // Create rules
            outlookWrapper.CreateRules();
            // Play sound to indicate success
            SystemSounds.Exclamation.Play();
        }
    }
}
