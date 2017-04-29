using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;
using System.Security.Principal;
using System.Diagnostics;

namespace Editor
{
    class Program
    {

        public static Dictionary<char, String> nmlsPathes = new Dictionary<char, string>();
        public static Dictionary<char, bool> nmlsDelDrives = new Dictionary<char, bool>();
        public static DirectoryEntry directoryEntry;

        static void Main(string[] args)
        {
            try
            {
                DirectoryEntry searchRoot = new DirectoryEntry("LDAP://rootDSE");
                Object domainDCs = searchRoot.Properties["defaultNamingContext"][0];
                searchRoot.Dispose();

                Console.WriteLine("Domain is: " + domainDCs);

                directoryEntry = new DirectoryEntry(args[0]);
                // System.Windows.Forms.MessageBox.Show(args[0]);

                if (directoryEntry == null)
                {
                    Environment.Exit(100);
                }
                else
                {
                    Console.WriteLine("Found directory object in ActiveDirectory: " + directoryEntry.Name);
                }

                HashSet<DirectoryEntry> ouPath = BuildOUPath(directoryEntry);

                foreach (DirectoryEntry elem in ouPath.Reverse())
                {
                    //System.Windows.Forms.MessageBox.Show(elem.Name);
                    ReadNmlsUncPathfromLDAP(elem, nmlsPathes);
                    ReadNmlsDelDriveFromLDAP(elem, nmlsDelDrives);
                }

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(true);
                Application.Run(new EditorProgram());
            }
            catch (Exception e)
            {
                Console.WriteLine(e.HResult + " " + e.Message);
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

        private static HashSet<DirectoryEntry> BuildOUPath(DirectoryEntry user)
        {
            HashSet<DirectoryEntry> result = new HashSet<DirectoryEntry>();

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
                    //System.Windows.Forms.MessageBox.Show(uncPath);
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
    }
}
