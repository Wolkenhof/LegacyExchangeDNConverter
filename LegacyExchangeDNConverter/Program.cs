using System.Diagnostics;
using CsvHelper;
using CsvHelper.Configuration.Attributes;
using System.Globalization;
using CsvHelper.Configuration;
using System.Text.RegularExpressions;
using LegacyExchangeDNConverter.Common;
using System.Diagnostics.CodeAnalysis;

namespace LegacyExchangeDNConverter
{
    internal class Program
    {
        public class User
        {
            [Name("Name")]
            public string Name { get; set; }
            [Name("LegacyExchangeDN")]
            public string LegacyExchangeDN { get; set; }
            [Name("ProxyAddresses")]
            public string ProxyAddresses { get; set; }
            [Name("ProxyAddresses2")]
            public string ProxyAddresses2 { get; set; }
        }
        
        public class User2
        {
            [Name("Name")]
            public string Name { get; set; }
            [Name("UserPrincipalName")]
            public string UserPrincipalName { get; set; }
        }

        private static bool _isDebug = false;

        public static string ConvertIMCEAEXToX500(string imceaex)
        {
            if (string.IsNullOrEmpty(imceaex))
            {
                throw new ArgumentNullException(nameof(imceaex), "IMCEAEX address cannot be null or empty.");
            }

            var result = imceaex;
            result = result.Replace("_", "/");
            result = result.Replace("+20", " ");
            result = result.Replace("+28", "(");
            result = result.Replace("+29", ")");
            result = result.Replace("+40", "@");
            result = result.Replace("+2E", ".");
            result = result.Replace("+2C", ",");
            result = result.Replace("+5F", "_");
            result = result.Replace("IMCEAEX-", "X500:");
            result = result.Split('@')[0];

            return result;
        }

        [DynamicDependency(DynamicallyAccessedMemberTypes.All, typeof(User))]
        [DynamicDependency(DynamicallyAccessedMemberTypes.All, typeof(CsvHelper.Configuration.DefaultClassMap<User>))]
        [DynamicDependency(DynamicallyAccessedMemberTypes.All, typeof(CsvHelper.Configuration.MemberMap<User, string>))]
        [DynamicDependency(DynamicallyAccessedMemberTypes.All, typeof(CsvHelper.Expressions.RecordManager))]
        [DynamicDependency(DynamicallyAccessedMemberTypes.All, typeof(CsvHelper.Expressions.RecordCreatorFactory))]
        [DynamicDependency(DynamicallyAccessedMemberTypes.All, typeof(CsvHelper.Expressions.RecordHydrator))]
        [DynamicDependency(DynamicallyAccessedMemberTypes.All, typeof(CsvHelper.Expressions.ExpressionManager))]
        [DynamicDependency(DynamicallyAccessedMemberTypes.All, typeof(CsvHelper.TypeConversion.StringConverter))]
        static void Main(string[] args)
        {

            var imceaexAddress = "IMCEAEX-_o=Voelkel_ou=Exchange+20Administrative+20Group+20+28FYDIBOHF23SPDLT+29_cn=Recipients_cn=Leistikow+2C+20Britta+20-+20Voelkel+20GmbHc7c@eurprd04.prod.outlook.com";
            var x500Address = ConvertIMCEAEXToX500(imceaexAddress);
            //DebugConsole.WriteLine("Converted IMCEAEX: " + x500Address);
            //DebugConsole.WriteLine("Example second:    x500:/o=ExchangeLabs/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)/cn=Recipients/cn=5e931242529445eea2b6c7b52051c052-Gruender");
            //Console.WriteLine();

            var convertPath = string.Empty;
            var writePath = string.Empty;
            var filterPath = string.Empty;
            var useConvert = false;
            var useWrite = false;
            var useFilter = false;
#if DEBUG
            _isDebug = true;
#endif

            if (args.Length == 0)
            {
                Console.WriteLine("Syntax: LegacyExchangeDNConverter.exe /convert:<Path\\To\\CSV> /write:<Path\\To\\CSV>");
                Environment.Exit(1);
            }

            foreach (var arg in args)
            {
                if (arg.StartsWith("/convert:", StringComparison.OrdinalIgnoreCase))
                {
                    convertPath = arg.Substring("/convert:".Length);
                    if (!File.Exists(convertPath))
                    {
                        DebugConsole.WriteLine("CSV Datei nicht gefunden: " + convertPath, ConsoleColor.Red);
                        Environment.Exit(1);
                    }

                    useConvert = true;
                }
                else if (arg.StartsWith("/write:", StringComparison.OrdinalIgnoreCase))
                {
                    writePath = arg.Substring("/write:".Length);
                    if (!File.Exists(writePath))
                    {
                        DebugConsole.WriteLine("CSV Datei nicht gefunden: " + writePath, ConsoleColor.Red);
                        Environment.Exit(1);
                    }

                    useWrite = true;
                }
                else if (arg.StartsWith("/filter:", StringComparison.OrdinalIgnoreCase))
                {
                    filterPath = arg.Substring("/filter:".Length);
                    if (!File.Exists(filterPath))
                    {
                        DebugConsole.WriteLine("CSV Datei nicht gefunden: " + filterPath, ConsoleColor.Red);
                        Environment.Exit(1);
                    }

                    useFilter = true;
                }
            }


            if (useConvert)
            {
                try
                {
                    // Import CSV
                    using var reader = new StreamReader(convertPath);
                    var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture)
                    {
                        HasHeaderRecord = true,
                        HeaderValidated = null,
                        MissingFieldFound = null
                    });
                    var records = csv.GetRecords<User>().ToList();

                    var converted = new List<User>();
                    foreach (var record in records)
                    {
                        DebugConsole.WriteLine("Converting User: " + record.Name);

                        // Get '/o'
                        var parts = record.LegacyExchangeDN.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
                        var organization = parts.FirstOrDefault(part => part.StartsWith("o=", StringComparison.OrdinalIgnoreCase));
                        if (!string.IsNullOrEmpty(organization))
                            organization = $"/{organization}";

                        // Set '/ou'
                        var organizationUnit = "";
                        if (record.LegacyExchangeDN.Contains("/ou=Exchange Administrative Group"))
                            organizationUnit = "/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)";
                        else if (record.LegacyExchangeDN.Contains("/ou=External"))
                            organizationUnit = "/ou=External (FYDIBOHF25SPDLT)";

                        // Set '/cn'
                        var name = record.Name;
                        name = name.Replace(",", "");
                        var commonName1 = "/cn=Recipients";
                        var commonName2 = $"/cn={name}";
                        var x500 = $"X500:{organization}{organizationUnit}{commonName1}{commonName2}";
                        DebugConsole.WriteLine("Created X500: " + x500);

                        // Create X500 + /cn:<ID>
                        // Use a regular expression to find the /cn=... part
                        var matches = Regex.Matches(record.LegacyExchangeDN, @"/cn=[^/]+");
                        string? match = null;
                        if (matches.Count == 0)
                        {
                            DebugConsole.WriteLine("No Common Name ID defined.", ConsoleColor.Yellow);
                        }
                        else
                        {
                            match = matches[^1].Value;
                        }

                        //var secondX500O = "/o=ExchangeLabs";
                        var x500WithCn = $"X500:{organization}{organizationUnit}{commonName1}{match}";
                        DebugConsole.WriteLine("Created X500: " + x500WithCn);

                        // Update LegacyExchangeDN
                        var legacyExchangeDn = record.LegacyExchangeDN;
                        try
                        {
                            // Find the index of the first '/' after "/o="
                            var startIndex = legacyExchangeDn.IndexOf("/o=", StringComparison.OrdinalIgnoreCase);
                            if (startIndex == -1)
                            {
                                DebugConsole.WriteLine("No Organization defined.", ConsoleColor.Yellow);
                            }
                            else
                            {
                                // Find the index of the next '/' after "/o="
                                var endIndex = legacyExchangeDn.IndexOf('/', startIndex + 1);
                                if (endIndex == -1)
                                {
                                    // If endIndex is -1, it means there are no more '/' after "/o=", so just remove from startIndex to end
                                    if (startIndex >= 0 && startIndex <= legacyExchangeDn.Length)
                                        legacyExchangeDn = legacyExchangeDn[..startIndex];
                                }
                                else
                                {
                                    // Remove the "/o=Example" part including the trailing '/'
                                    legacyExchangeDn = legacyExchangeDn[..startIndex] +
                                                       legacyExchangeDn[endIndex..];
                                }

                                //DebugConsole.WriteLine("Updated LegacyExchangeDN: " + legacyExchangeDn);
                            }
                        }
                        catch (Exception ex)
                        {
                            DebugConsole.WriteLine("An error has occurred: " + ex.Message, ConsoleColor.Red);
                            if (_isDebug) Debugger.Break();
                        }

                        converted.Add(new User()
                        {
                            Name = record.Name,
                            ProxyAddresses = x500,
                            LegacyExchangeDN = legacyExchangeDn,
                            ProxyAddresses2 = x500WithCn
                        });

                        Console.WriteLine();
                    }

                    // Write to file
                    var destination = Path.Combine(Path.GetDirectoryName(convertPath)!, $"{Path.GetFileNameWithoutExtension(convertPath)}.updated.csv");

                    using var writer = new StreamWriter(destination);
                    using var newCsv = new CsvWriter(writer, CultureInfo.InvariantCulture);
                    newCsv.WriteRecords(converted);
                    DebugConsole.WriteLine("Converting completed!", ConsoleColor.Green);
                }
                catch (Exception ex)
                {
                    DebugConsole.WriteLine("An error has occurred: " + ex.Message, ConsoleColor.Red);
                    if (_isDebug) Debugger.Break();
                    Environment.Exit(1);
                }

            }
            else if (useWrite)
            {
                // Import CSV
                using var reader = new StreamReader(writePath);
                var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture)
                {
                    HasHeaderRecord = true,
                    HeaderValidated = null,
                    MissingFieldFound = null
                });
                var records = csv.GetRecords<User>().ToList();

                foreach (var record in records)
                {
                    DebugConsole.WriteLine("Modifiziere Attribute für Benutzer: " + record.Name);
                    //ADManager.UpdateLegacyExchangeDN(record.Name, record.LegacyExchangeDN);
                    ADManager.UpdateProxyAddresses(record.Name, record.ProxyAddresses, record.ProxyAddresses2);
                }
            }
            else if (useFilter)
            {
                var a = File.ReadAllLines("C:\\Users\\jgu\\Desktop\\LegacyExchangeDN\\missing.txt");
                using var reader = new StreamReader(filterPath);
                var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture)
                {
                    HasHeaderRecord = true,
                    HeaderValidated = null,
                    MissingFieldFound = null
                });
                var records = csv.GetRecords<User2>().ToList();
                var converted = new List<User2>();
                foreach (var record in records)
                {
                    if (a.Contains(record.Name))
                    {
                        converted.Add(new User2()
                        {
                            UserPrincipalName = record.UserPrincipalName,
                            Name = record.Name
                        });
                    }
                }

                // Write to file
                var destination = Path.Combine(Path.GetDirectoryName(filterPath)!, $"{Path.GetFileNameWithoutExtension(filterPath)}.updated.csv");

                using var writer = new StreamWriter(destination);
                using var newCsv = new CsvWriter(writer, CultureInfo.InvariantCulture);
                newCsv.WriteRecords(converted);
                DebugConsole.WriteLine("Converting completed!", ConsoleColor.Green);
            }
            else
            {
                DebugConsole.WriteLine("Invalid command!", ConsoleColor.Red);
            }
        }
    }
}
