using System;
using System.Data;
using System.Management.Automation;
using System.Management.Automation.Remoting;
using System.Management.Automation.Runspaces;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Security;
using System.Text;
using System.IO;

namespace MeetingAnalyzer
{
    class Program
    {
        static void Main(string[] args)
        {
            string strAdminName = "";
            string strPwd = "";
            string strMailbox = "";
            string strSubject = "";
            string strGOID = "";
            bool bUseSubject = true;
            DataSet dsMsgs = new DataSet("Messages");

            Console.WriteLine("");
            Console.WriteLine("================");
            Console.WriteLine("Meeting Analyzer");
            Console.WriteLine("================");
            Console.WriteLine("Creates a timeline of a meeting and reports any problems found.\r\n");
            
            // get the tenant admin
            Console.Write("Enter the tenant admin name (eg. admin@tailspintoys.onmicrosoft.com): ");
            strAdminName = Console.ReadLine();
            // get the password
            Console.Write("Enter the password for {0}: ", strAdminName);
            // use below while loop to mask the password while reading it in
            bool bEnter = true;
            int iPwdChars = 0;
            while (bEnter)
            {
                ConsoleKeyInfo ckiKey = Console.ReadKey(true);
                if (ckiKey.Key == ConsoleKey.Enter)
                {
                    bEnter = false;
                }
                else if (ckiKey.Key == ConsoleKey.Backspace)
                {
                    if (strPwd.Length >= 1)
                    {
                        int oldLength = strPwd.Length;
                        strPwd = strPwd.Substring(0, oldLength - 1);
                        Console.Write("\b \b");
                    }
                }
                else
                {
                    strPwd = strPwd + ckiKey.KeyChar.ToString();
                    iPwdChars++;
                    Console.Write('*');
                }
            }
            Console.WriteLine("");
            // have to make the password secure for the connection
            SecureString secPwd = new SecureString();
            foreach(char c in strPwd)
            {
                secPwd.AppendChar(c);
            }
            // get rid of the textual password as soon as we have the secure password.
            strPwd = "";
            strPwd = null;

            // Now we have our username+password creds we can use for the connection
            PSCredential psCred = new PSCredential(strAdminName, secPwd);

            // Get rid of admin username now that the cred object is created
            strAdminName = "";
            strAdminName = null;

            // Now make the connection object for the service
            Uri uriPS = new Uri("https://outlook.office365.com/powershell-liveid/"); // O365 - https://outlook.office365.com/powershell-liveid/ (from https://technet.microsoft.com/en-us/library/jj984289(v=exchg.160).aspx)
            string strShellUri = "http://schemas.microsoft.com/powershell/Microsoft.Exchange";  
            WSManConnectionInfo connectionInfo = new WSManConnectionInfo(uriPS, strShellUri, psCred);
            connectionInfo.AuthenticationMechanism = AuthenticationMechanism.Basic;
            connectionInfo.MaximumConnectionRedirectionCount = 5;

            Collection<PSObject> pshResults = new Collection<PSObject>(); // results collection object from running commands

            // Now do the connection and run PS commands
            using (Runspace rs = RunspaceFactory.CreateRunspace(connectionInfo))
            {
                PowerShell psh = PowerShell.Create();
                Console.WriteLine("\r\nAttempting to connect to the service.");
                try
                {
                    rs.Open();
                    psh.Runspace = rs;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Could not open the remote powershell session."); 
                    Console.WriteLine("Error: " + ex.Message.ToString());
                    goto Exit;
                }
                
                Console.WriteLine("Successfully connected.\r\n");

                //Get rid of the secure password, don't need it anymore.
                secPwd.Dispose();
                secPwd = null;

            NextMtg:

                // Now get user and meeting info from the user
                Console.Write("Enter the SMTP address of the user for the meeting to analyze: ");
                strMailbox = Console.ReadLine();
                Timeline.m_strMbx = strMailbox; // For printing out to screen later
                Utils.m_MBX = strMailbox;
                Console.Write("Enter the Subject of the meeting to analyze (leave blank to enter a Meeting ID instead): ");
                strSubject = Console.ReadLine();
                if (string.IsNullOrEmpty(strSubject))
                {
                    Console.Write("Enter the Meeting ID (Global Object ID) of the problem meeting: ");
                    strGOID = Console.ReadLine();

                    if(!(string.IsNullOrEmpty(strMailbox)) && !(string.IsNullOrEmpty(strGOID)))
                    {
                        Utils.CreateFile(strMailbox, strGOID);
                        bUseSubject = false;
                    }
                    else
                    {
                        string strYN = "no";
                        Console.Write("\r\nSMTP address and Subject or Global Object ID must have values. Do you want to try again (yes/no)? ");
                        strYN = Console.ReadLine().ToLower();
                        Console.WriteLine();
                        if (strYN.StartsWith("y"))
                        {
                            strMailbox = "";
                            Timeline.m_strMbx = "";
                            Utils.m_MBX = "";
                            strSubject = "";
                            strGOID = "";
                            goto NextMtg;
                        }
                        else
                        {
                            goto Exit;
                        }
                    }
                }
                else
                {
                    Utils.m_Subject = strSubject;
                    if (!(string.IsNullOrEmpty(strMailbox)) && !(string.IsNullOrEmpty(strSubject)))
                    {
                        Utils.CreateFile(strMailbox, strSubject);
                    }
                    else
                    {
                        string strYN = "no";
                        Console.Write("\r\nSMTP address and Subject or Global Object ID must have values. Do you want to try again (yes/no)? ");
                        strYN = Console.ReadLine().ToLower();
                        Console.WriteLine();
                        if (strYN.StartsWith("y"))
                        {
                            strMailbox = "";
                            Timeline.m_strMbx = "";
                            Utils.m_MBX = "";
                            strSubject = "";
                            strGOID = "";
                            goto NextMtg;
                        }
                        else
                        {
                            goto Exit;
                        }
                    }
                }

                Console.WriteLine("\r\nRunning command to retreive the meeting data...");

                // Run Get-CalendarDiagnosticObjects to get the version history data
                // Get-CalendarDiagnosticObjects -Identity <user smtp address> -Subject <meeting subject>  -OutputProperties Custom -CustomPropertyNames ItemClass,NormalizedSubject...
                try
                {
                    psh.AddCommand("Get-CalendarDiagnosticObjects");
                    psh.AddParameter("-Identity", strMailbox);
                    if (bUseSubject)
                    {
                        psh.AddParameter("-Subject", strSubject);
                    }
                    else
                    {
                        psh.AddParameter("-MeetingId", strGOID);
                    }
                    psh.AddParameter("-OutputProperties", "Custom");
                    psh.AddParameter("-CustomPropertyNames", Utils.rgstrPropsToGet);

                    pshResults = psh.Invoke();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("\r\nUnable to process the Get-CalendarDiagnosticObjects command.");
                    Console.WriteLine("Error: " + ex.Message.ToString());
                    goto Exit;
                }

                if (pshResults.Count > 0)
                {
                    MsgData md = new MsgData();
                    Console.WriteLine("Successfully retreived the meeting data.\r\n");

                    Timeline.m_iNumMsgs = pshResults.Count;
                    Console.Write("Importing data from messages.");

                    foreach (PSObject message in pshResults)
                    {
                        MsgData.PopulateDataSet(message);
                        Console.Write(".");
                    }
                    Console.WriteLine("\r\nData import completed.\r\n");
                }
                else
                {
                    Console.WriteLine("No meeting data was retreived. Check the user and meeting information and try again.\r\n");
                    psh.Commands.Clear();
                    pshResults.Clear();
                    Utils.Reset();
                    strGOID = "";
                    MsgData.m_GOID = "";
                    File.Delete(Utils.m_FilePath);
                    goto NextMtg;
                }
                // we got the data imported - now we can sort by the Modified Time
                dsMsgs = MsgData.dsSort(MsgData.msgDataSet, "OrigModTime");
                
                // Create the timeline with the sorted messages
                Timeline.CreateTimeline(dsMsgs);

                Utils.CreateCSVFile(Utils.m_MBX, Utils.m_Subject);
                Utils.WriteMsgData(dsMsgs);

                Console.WriteLine("Timeline output written to:             {0}", Utils.m_FilePath);
                Console.WriteLine("Calendar item property data written to: {0}", Utils.m_CSVPath);

                string strYesNo = "no";
                //strYesNo = "no";
                Console.Write("\r\nAnalyze data for another meeting (yes/no)? ");
                strYesNo = Console.ReadLine().ToLower();
                Console.WriteLine();
                if (strYesNo.StartsWith("y"))
                {
                    dsMsgs.Clear();
                    psh.Commands.Clear();
                    pshResults.Clear();
                    Utils.Reset();
                    strGOID = "";
                    MsgData.m_GOID = "";
                    Console.WriteLine("==================================================");
                    Console.WriteLine();
                    goto NextMtg;
                }
                else
                {
                    // clear out stuff
                    dsMsgs.Clear();
                    psh.Commands.Clear();
                    pshResults.Clear();
                    Utils.Reset();
                    strGOID = "";
                    MsgData.m_GOID = "";

                    // close out the runspace since we're done.
                    rs.Close();
                    rs.Dispose();
                }
            }

            Exit:
            // Exit the app...
            Console.Write("\r\nExiting the program.");
            //Console.ReadLine();
        }
    }
}
