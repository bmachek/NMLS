using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;
using System.Security.Principal;
using System.Diagnostics;

namespace Logon
{
    class ProgramLogon
    {

        static void Main(string[] args)
        {
            DirectoryEntry searchRoot = new DirectoryEntry("LDAP://rootDSE");
            Object domainDCs = searchRoot.Properties["defaultNamingContext"][0];
            searchRoot.Dispose();

            Console.WriteLine("Domain is: " + domainDCs);

            WindowsIdentity identity = WindowsIdentity.GetCurrent();
            String username = identity.Name.Substring(identity.Name.IndexOf(@"\") + 1);


            DirectoryEntry userEntry = FindUserInActiveDirectory(domainDCs, username);

            if (userEntry == null)
            {
                Environment.Exit(100);
            }
            else
            {
                Console.WriteLine("Found user object in ActiveDirectory: " + userEntry.Name);
            }

            HashSet<DirectoryEntry> path = BuildOUPath(userEntry);

            Dictionary<char, String> nmlsPathes = new Dictionary<char, string>();
            Dictionary<char, bool> nmlsDelDrives = new Dictionary<char, bool>();

            foreach (DirectoryEntry elem in path.Reverse())
            {
                ReadNmlsUncPathfromLDAP(elem, nmlsPathes);
                ReadNmlsDelDriveFromLDAP(elem, nmlsDelDrives);
            }

            foreach (KeyValuePair<char, bool> item in nmlsDelDrives)
            {
                if (item.Value)
                {
                    // Console.WriteLine("Deleting drive " + item.Key);
                    MapNetworkDriveDisconnect(item.Key.ToString());
                }
            }

            foreach (KeyValuePair<char, string> item in nmlsPathes)
            {
                if (item.Value != null && item.Value != "")
                {
                    Console.WriteLine("Connecting drive " + item.Key + " with " + item.Value);
                    MapNetworkDriveConnect(item.Key.ToString(), item.Value);
                }
            }
        }

        private static void ReadNmlsUncPathfromLDAP(DirectoryEntry userEntry, Dictionary<char, String> result)
        {
            for (int i = 1; i <= 26; i++)
            {

                char driveLetter = Convert.ToChar(i + 64);

                object uncPathObj = userEntry.Properties["nmlsNetworkDrive" + driveLetter].Value;
                string uncPath = (String)uncPathObj;

                if (uncPathObj != null)
                {

                    if (result.ContainsKey(driveLetter))
                    {

                        Console.WriteLine("Overriding previous UNC path assignment: " + result[driveLetter]);
                        result.Remove(driveLetter);
                    }
                    result.Add(driveLetter, uncPath);
                }
            }
        }

        private static void ReadNmlsDelDriveFromLDAP(DirectoryEntry userEntry, Dictionary<char, bool> result)
        {
            for (int i = 1; i <= 26; i++)
            {

                char driveLetter = Convert.ToChar(i + 64);

                object driveDelObject = userEntry.Properties["nmlsNetworkDriveDel" + driveLetter].Value;

                if (driveDelObject != null)
                {
                    if (result.ContainsKey(driveLetter))
                    {
                        result.Remove(driveLetter);
                    }
                    result.Add(driveLetter, (bool)driveDelObject);
                }
            }
        }


        private static DirectoryEntry FindUserInActiveDirectory(object domainDCs, String username)
        {
            DirectorySearcher searcher = new DirectorySearcher(domainDCs.ToString());
            searcher.Filter = "(&(sAMAccountName=" + username + "))";

            SearchResultCollection result = searcher.FindAll();

            if (result.Count == 1)
            {
                return new DirectoryEntry(result[0].Path);
            }
            else if (result.Count == 0)
            {
                return null;
            }
            else
            {
                return null;
            }
        }

        private static void MapNetworkDriveDisconnect(string drive)
        {

            Console.WriteLine("Trying to disconnect " + drive);

            Process p = new Process();
            p.StartInfo.FileName = "net";
            p.StartInfo.Arguments = string.Format("use {0}: /DELETE ", drive);
            p.StartInfo.UseShellExecute = false;
            p.Start();
            p.WaitForExit();
        }

        private static void MapNetworkDriveConnect(string drive, string server)
        {
            Console.WriteLine("Trying to connect " + drive + " with " + server);

            Process p = new Process();
            p.StartInfo.FileName = "net";
            p.StartInfo.Arguments = string.Format("use {0}: {1}", drive, server);
            p.StartInfo.UseShellExecute = false;
            p.Start();
            p.WaitForExit();
        }

        private static HashSet<DirectoryEntry> BuildOUPath(DirectoryEntry user)
        {
            HashSet<DirectoryEntry> result = new HashSet<DirectoryEntry>();
            result.Add(user);

            DirectoryEntry parent = user.Parent;
            while (true)
            {
                if (parent.SchemaClassName.Equals("organizationalUnit"))
                {
                    // Console.WriteLine(parent.Path);
                    result.Add(parent);
                    parent = parent.Parent;
                }
                else
                {
                    break;
                }
            }

            return result;
        }
    }
}
