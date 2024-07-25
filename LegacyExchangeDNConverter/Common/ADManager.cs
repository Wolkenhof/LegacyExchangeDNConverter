using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;

namespace LegacyExchangeDNConverter.Common
{
    public class ADManager
    {
        public static void UpdateProxyAddresses(string name, string newProxyAddresses, string newProxyAddresses2)
        {
            try
            {
                var deviceDomain = Environment.UserDomainName;
                using var context = new PrincipalContext(ContextType.Domain, deviceDomain);
                using var searcher = new PrincipalSearcher(new UserPrincipal(context) { Name = name});
                var result = searcher.FindOne();
                if (result != null)
                {
                    if (result.GetUnderlyingObject() is DirectoryEntry de)
                    {
                        de.Properties["proxyAddresses"].Add(newProxyAddresses);
                        de.Properties["proxyAddresses"].Add(newProxyAddresses2);
                        de.CommitChanges();
                        var errorMessage = "Attribut 'proxyAddresses' wurden erfolgreich geändert!";
                        DebugConsole.WriteLine(errorMessage, ConsoleColor.Green);
                    }
                }
                else
                {
                    var errorMessage = "Benutzer nicht gefunden!";
                    DebugConsole.WriteLine(errorMessage, ConsoleColor.Red);
                }
            }
            catch (Exception ex)
            {
                var errorMessage = "Fehler: " + ex.Message;
                DebugConsole.WriteLine(errorMessage, ConsoleColor.Red);
            }
        }

        public static void UpdateLegacyExchangeDN(string name, string newLegacyExchangeDN)
        {
            try
            {
                var deviceDomain = Environment.UserDomainName;
                using var context = new PrincipalContext(ContextType.Domain, deviceDomain);
                using var searcher = new PrincipalSearcher(new UserPrincipal(context) { Name = name });
                var result = searcher.FindOne();
                if (result != null)
                {
                    if (result.GetUnderlyingObject() is DirectoryEntry de)
                    {
                        de.Properties["legacyExchangeDN"].Clear();
                        de.Properties["legacyExchangeDN"].Add(newLegacyExchangeDN);
                        de.CommitChanges();
                        var errorMessage = "Attribut 'legacyExchangeDN' wurden erfolgreich geändert!";
                        DebugConsole.WriteLine(errorMessage, ConsoleColor.Green);
                    }
                }
                else
                {
                    var errorMessage = "Benutzer nicht gefunden!";
                    DebugConsole.WriteLine(errorMessage, ConsoleColor.Red);
                }
            }
            catch (Exception ex)
            {
                var errorMessage = "Fehler: " + ex.Message;
                DebugConsole.WriteLine(errorMessage, ConsoleColor.Red);
            }
        }
    }
}
