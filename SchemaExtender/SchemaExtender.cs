using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;

namespace NMLS
{
    class Program
    {

        static void Main(string[] args)
        {
            ActiveDirectorySchema schema = ActiveDirectorySchema.GetCurrentSchema();

            DirectoryEntry searchRoot = new DirectoryEntry("LDAP://rootDSE");
            Object domainDCs = searchRoot.Properties["defaultNamingContext"][0];
            searchRoot.Dispose();

            Console.WriteLine("Processing for domain: " + domainDCs);

            for (int i = 1; i <= 26; i++)
            {
                char driveLetter = Convert.ToChar(i + 64);
                Console.WriteLine("Processing drive: " + driveLetter);
                createNetworkDriveAttribute(domainDCs, driveLetter, i);

            }

            schema.RefreshSchema();

            createSchemaClass(domainDCs);

            schema.RefreshSchema();

            CreateDisplaySpecifiers(domainDCs);

            searchRoot.Dispose();

        }


        private static void createNetworkDriveAttribute(Object domainDCs, char driveLetter, int attributeId)
        {
            try
            {
                DirectoryEntry schemaPartition = new DirectoryEntry("LDAP://cn=Schema,cn=Configuration," + domainDCs);

                if (!DirectoryEntry.Exists("LDAP://CN=nmlsNetworkDrive" + driveLetter + ",cn=Schema,cn=Configuration," + domainDCs))
                {
                    DirectoryEntry newNetworkDrive = schemaPartition.Children.Add("CN=nmlsNetworkDrive" + driveLetter, "attributeSchema");
                    newNetworkDrive.Properties["objectClass"].Value = "attributeSchema";
                    newNetworkDrive.Properties["attributeId"].Value = "1.2.840.113556.1.8000.2554.61823.22188.26389.18272.44952.1599828.8363415.1." + attributeId;
                    newNetworkDrive.Properties["attributeSyntax"].Value = "2.5.5.12";
                    newNetworkDrive.Properties["oMSyntax"].Value = "64";
                    newNetworkDrive.Properties["isSingleValued"].Value = "TRUE";
                    newNetworkDrive.Properties["rangeLower"].Value = "1";
                    newNetworkDrive.Properties["rangeUpper"].Value = "50";


                    newNetworkDrive.CommitChanges();

                }
                else
                {
                    Console.WriteLine("networkDrive" + driveLetter + " already exists");

                }

                if (!DirectoryEntry.Exists("LDAP://CN=nmlsNetworkDriveDel" + driveLetter + ",cn=Schema,cn=Configuration," + domainDCs))
                {
                    DirectoryEntry newNetworkDrive = schemaPartition.Children.Add("CN=nmlsNetworkDriveDel" + driveLetter, "attributeSchema");
                    newNetworkDrive.Properties["objectClass"].Value = "attributeSchema";
                    newNetworkDrive.Properties["attributeId"].Value = "1.2.840.113556.1.8000.2554.61823.22188.26389.18272.44952.1599828.8363415.2." + attributeId;
                    newNetworkDrive.Properties["attributeSyntax"].Value = "2.5.5.8";
                    newNetworkDrive.Properties["oMSyntax"].Value = "1";
                    newNetworkDrive.Properties["isSingleValued"].Value = "TRUE";
                    newNetworkDrive.Properties["rangeLower"].Value = "1";
                    newNetworkDrive.Properties["rangeUpper"].Value = "50";


                    newNetworkDrive.CommitChanges();

                }
                else
                {
                    Console.WriteLine("networkDriveDel" + driveLetter + " already exists");

                }

            }
            catch (Exception e)
            {
                Console.WriteLine("Fehler bei Laufwerk: " + driveLetter);
                Console.WriteLine(e.ToString());
            }
        }

        private static void createSchemaClass(object domainDCs)
        {
            try
            {
                if (!DirectoryEntry.Exists("LDAP://CN=nmlsNetworkDriveClass,cn=Schema,cn=Configuration," + domainDCs))
                {

                    DirectoryContext ctx = new DirectoryContext(DirectoryContextType.Forest);

                    ActiveDirectorySchemaClass schemaClass = new ActiveDirectorySchemaClass(ctx, "nmlsNetworkDriveClass");
                    schemaClass.Oid = "1.2.840.113556.1.8000.2554.61823.22188.26389.18272.44952.1599828.8363415.3.1";
                    schemaClass.SubClassOf = new ActiveDirectorySchemaClass(ctx, "top");
                    schemaClass.OptionalProperties.AddRange(new ActiveDirectorySchemaProperty[] {
                        ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveA"),
                        ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveB"),
                        ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveC"),
                        ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveD"),
                        ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveE"),
                        ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveF"),
                        ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveG"),
                        ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveH"),
                        ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveI"),
                        ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveJ"),
                        ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveK"),
                        ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveL"),
                        ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveM"),
                        ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveN"),
                        ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveO"),
                        ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveP"),
                        ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveQ"),
                        ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveR"),
                        ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveS"),
                        ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveT"),
                        ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveU"),
                        ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveV"),
                        ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveW"),
                        ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveX"),
                        ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveY"),
                        ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveZ"),
                       ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveDelA"),
                       ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveDelB"),
                       ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveDelC"),
                       ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveDelD"),
                       ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveDelE"),
                       ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveDelF"),
                       ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveDelG"),
                       ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveDelH"),
                       ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveDelI"),
                       ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveDelJ"),
                       ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveDelK"),
                       ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveDelL"),
                       ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveDelM"),
                       ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveDelN"),
                       ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveDelO"),
                       ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveDelP"),
                       ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveDelQ"),
                       ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveDelR"),
                       ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveDelS"),
                       ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveDelT"),
                       ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveDelU"),
                       ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveDelV"),
                       ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveDelW"),
                       ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveDelX"),
                       ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveDelY"),
                       ActiveDirectorySchemaProperty.FindByName(ctx, "nmlsNetworkDriveDelZ")
                    });

                    schemaClass.Type = System.DirectoryServices.ActiveDirectory.SchemaClassType.Auxiliary;

                    schemaClass.Save();

                    //schemaClass 
                }
                else
                {
                    Console.WriteLine("Class already exists");
                }

                DirectoryEntry userClass = new DirectoryEntry("LDAP://cn=User,cn=Schema,cn=Configuration," + domainDCs);

                if (!userClass.Properties["auxiliaryClass"].Contains("nmlsNetworkDriveClass"))
                {
                    userClass.Properties["auxiliaryClass"].Add("nmlsNetworkDriveClass");
                    userClass.CommitChanges();

                }
                else
                {
                    Console.WriteLine("User class already has auxiliary class nmlsNetworkDriveClass");
                }


                // Add NMLS class also OU objects.

                DirectoryEntry ouClass = new DirectoryEntry("LDAP://cn=Organizational-Unit,cn=Schema,cn=Configuration," + domainDCs);
                if (!ouClass.Properties["auxiliaryClass"].Contains("nmlsNetworkDriveClass"))
                {
                    ouClass.Properties["auxiliaryClass"].Add("nmlsNetworkDriveClass");
                    ouClass.CommitChanges();
                }
                else
                {
                    Console.WriteLine("OU class already has auxiliary class nmlsNetworkDriveClass");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
        }

        private static void CreateDisplaySpecifiers(Object dc)
        {
            DirectoryEntry germanUserSpecifiers = new DirectoryEntry("LDAP://CN=user-Display,CN=407,CN=DisplaySpecifiers,CN=Configuration," + dc);
            germanUserSpecifiers.Properties["adminContextMenu"].Add("50,NMLS,\\\\" + System.DirectoryServices.ActiveDirectory.Domain.GetCurrentDomain().ToString() + "\\NETLOGON\\NMLS-Attribute Editor.exe");
            germanUserSpecifiers.CommitChanges();

            DirectoryEntry englishUserSpecifiers = new DirectoryEntry("LDAP://CN=user-Display,CN=409,CN=DisplaySpecifiers,CN=Configuration," + dc);
            englishUserSpecifiers.Properties["adminContextMenu"].Add("50,NMLS,\\\\" + System.DirectoryServices.ActiveDirectory.Domain.GetCurrentDomain().ToString() + "\\NETLOGON\\NMLS-Attribute Editor.exe");
            englishUserSpecifiers.CommitChanges();

            DirectoryEntry germanOuSpecifiers = new DirectoryEntry("LDAP://CN=organizationalUnit-Display,CN=407,CN=DisplaySpecifiers,CN=Configuration," + dc);
            germanOuSpecifiers.Properties["adminContextMenu"].Add("50,NMLS,\\\\" + System.DirectoryServices.ActiveDirectory.Domain.GetCurrentDomain().ToString() + "\\NETLOGON\\NMLS-Attribute Editor.exe");
            germanOuSpecifiers.CommitChanges();

            DirectoryEntry englishOuSpecifiers = new DirectoryEntry("LDAP://CN=organizationalUnit-Display,CN=409,CN=DisplaySpecifiers,CN=Configuration," + dc);
            englishOuSpecifiers.Properties["adminContextMenu"].Add("50,NMLS,\\\\" + System.DirectoryServices.ActiveDirectory.Domain.GetCurrentDomain().ToString() + "\\NETLOGON\\NMLS-Attribute Editor.exe");
            englishOuSpecifiers.CommitChanges();
        }

        private static String convertDCtoDNS(Object dc)
        {
            String result = (string)dc;

            result.Replace(",", ".");
            result.Replace("DC=", "");

            return result;
        }
    }

}
