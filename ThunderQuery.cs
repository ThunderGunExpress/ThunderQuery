using System;
using System.Collections.Generic;
using System.Text;
using System.Management;
using System.IO;
using System.Linq;
using CommandLine;
using NetTools;
using System.Net.NetworkInformation;
using System.Threading;

public class ThunderQ
{
    //Command Line Interface
    public class Options
    {
        [Option('t', "targets", Required = false, Default = "127.0.0.1", HelpText = "Targets to poll. Comma separated or hyphen range")]
        public string TargetString { get; set; }

        [Option('p', "pollamount", Default = 10, Required = false, HelpText = "Amount of times tool will poll targets every 30 seconds. Default 10 times - 10 * 30 seconds = 5 minutes.")]
        public int PollAmount { get; set; }

        [Option('s', "skipdetection", Default = false, Required = false, HelpText = "Tool will ping systems before attempting to enumerate via WMI. Use this switch to skip live detection.")]
        public bool SkipDetection { get; set; }

        [Option('a', "autoremove", Default = true, Required = false, HelpText = "Will automatically remove systems which fail WMI query.")]
        public bool AutoRemove { get; set; }

        [Option('v', "verbose", Default = false, Required = false, HelpText = "Be verbose")]
        public bool Verbose { get; set; }
        public string GetUsage()
        {
            var usage = new StringBuilder();
            usage.AppendLine("ThunderQuery");
            usage.AppendLine("-t or --targets (Default - 127.0.0.1)");
            usage.AppendLine("-p or --pollamount (Default - 10)");
            usage.AppendLine("-v or --verbose (Default - false)");
            usage.AppendLine("Examples");
            usage.AppendLine("ThunderQuery.exe -t 192.168.1.10-20 -p 15 -v");
            usage.AppendLine("ThunderQuery.exe -t 192.168.1.10,192.168.1.11,192.168.1.12 -p 20 -v");
            return usage.ToString();
        }
    }

    //Ping Host used for live target detection
    //https://stackoverflow.com/questions/11800958/using-ping-in-c-sharp
    static bool PingHost(string nameOrAddress)
    {
        bool pingable = false;
        Ping pinger = null;

        try
        {
            pinger = new Ping();
            PingReply reply = pinger.Send(nameOrAddress);
            pingable = reply.Status == IPStatus.Success;
        }
        catch (PingException)
        {
            // Discard PingExceptions and return false;
        }
        finally
        {
            if (pinger != null)
            {
                pinger.Dispose();
            }
        }

        return pingable;
    }

    //Management Scope Connection Function. Handles connection to WMI - local or remote.
    static ManagementScope ScopeConnect(string sConnectionString, int iTimeout)
    {
        try
        {            
            ManagementScope msScope = new ManagementScope(sConnectionString);
            msScope.Options.Timeout = TimeSpan.FromSeconds(iTimeout);
            msScope.Connect();
            return msScope;
        }
        catch
        {            
            return null;
        }
    }

    //Generate profile CSV of targets. Includes IP, hostname, domain, loggedon user, and domainrole
    static void GenerateTargetCSV(ManagementScope scope, CsvFileWriter writer)
    {
        try
        {
            scope.Options.Timeout = TimeSpan.FromSeconds(1);
            scope.Connect();
            if (!scope.IsConnected)
            {
                Console.WriteLine("Error connecting to remote machine");                
                return;
            }

            int iInterfaceIndex = 0;
            string sDomainRole = null;
            string sDNSHostname = null;
            string sUsername = null;
            string sDomain = null;
            String[] sIPAddress = null;
            
            //Find the Interface Index with the default route
            ObjectQuery oqRoute = new ObjectQuery("Select InterfaceIndex from Win32_IP4RouteTable WHERE destination = '0.0.0.0'");            
            ManagementObjectSearcher mgmtObjSearcherRoute = new ManagementObjectSearcher(scope, oqRoute);
            ManagementObjectCollection colRoute = mgmtObjSearcherRoute.Get();

            foreach (ManagementObject objRoute in colRoute)
                iInterfaceIndex = (int)objRoute["InterfaceIndex"];           

            //Find the IP address of the interface
            ObjectQuery oqIPAddr = new ObjectQuery("Select IPAddress from Win32_NetworkAdapterConfiguration WHERE InterfaceIndex = " + iInterfaceIndex);            
            ManagementObjectSearcher mgmtObjSearcherIP = new ManagementObjectSearcher(scope, oqIPAddr);
            ManagementObjectCollection colIP = mgmtObjSearcherIP.Get();

            foreach (ManagementObject objIP in colIP)
                sIPAddress = (String[])objIP["IPAddress"];            

            //Grab profile information - Hostname, Domain, LoggedOn User, and DomainRole
            ObjectQuery oqCompSys = new ObjectQuery("SELECT DNSHostName, Domain, Username, DomainRole FROM Win32_ComputerSystem");            
            ManagementObjectSearcher mgmtObjSearcher = new ManagementObjectSearcher(scope, oqCompSys);
            ManagementObjectCollection colCompSys = mgmtObjSearcher.Get();

            foreach (ManagementObject objSysInfo in colCompSys)
            {
                switch (Convert.ToInt32(objSysInfo["DomainRole"]))
                {
                    case 0:
                        sDomainRole = "StandAlone_Wks";
                        break;
                    case 1:
                        sDomainRole = "DomainMember_Wks";
                        break;
                    case 2:
                        sDomainRole = "StandAlone_Svr";
                        break;
                    case 3:
                        sDomainRole = "DomainMember_Svr";
                        break;
                    case 4:
                        sDomainRole = "BackupDC";
                        break;
                    case 5:
                        sDomainRole = "PrimaryDC";
                        break;
                    default:
                        sDomainRole = "Dunno";
                        break;
                }
                sDNSHostname = (string)objSysInfo["DNSHostName"];
                sDomain = (string)objSysInfo["Domain"];
                sUsername = (string)objSysInfo["Username"];
            }
            //Craft CSV for Profile
            CsvRow row = new CsvRow();
            row.Add(sDNSHostname);
            row.Add(sDomain);
            if (sUsername == null)
                row.Add("NoLoggedInUser");
            else
                row.Add(sUsername);
            row.Add(sDomainRole);
            row.Add(sIPAddress[0]);            
            writer.WriteRow(row);
            row.Clear();
            return;
        }
        catch (Exception e)
        {
            Console.WriteLine("{0} Expection Caught.", e);
        }
    }

    //Netstat WMI Function. Cycles to identify new TCP connections and exports it to CSV.
    static void CollectNetworkConnections(ManagementScope scope, CsvFileWriter writer)
    {
        try
        {
            scope.Options.Timeout = TimeSpan.FromSeconds(1);
            scope.Connect();
            if (!scope.IsConnected)
            {
                Console.WriteLine("Error connecting to remote machine");
                return;
            }

            ObjectQuery oqConns = new ObjectQuery("SELECT LocalAddress,LocalPort,RemoteAddress,RemotePort FROM MSFT_NetTCPConnection WHERE state = 5 AND RemoteAddress != '127.0.0.1' AND RemoteAddress != '::1' AND (RemotePort < 49152 OR LocalPort < 49152)");
            ManagementObjectSearcher mgmtObjSearcherConns = new ManagementObjectSearcher(scope, oqConns);
            ManagementObjectCollection colConns = mgmtObjSearcherConns.Get();

            CsvRow row = new CsvRow();

            foreach (ManagementObject objConns in colConns)
            {

                row.Add((string)objConns["LocalAddress"]);
                row.Add(objConns["LocalPort"].ToString());
                row.Add((string)objConns["RemoteAddress"]);
                row.Add(objConns["RemotePort"].ToString());
                writer.WriteRow(row);
                row.Clear();
            }
            return;
        }
        catch (Exception e)
        {
            Console.WriteLine("{0} Expection Caught.", e);
        }

    }

    //Function to Parse Target Input. Can take a range of IP addresses: 192.168.1.50-100 or comma separated targets 192.168.1.50,192.168.1.51
    static List<string> TargetParsing(string TargetString)
    {
        if (TargetString.Contains(","))
            return TargetString.Split(',').ToList();
        else if (TargetString.Contains("-"))
        {
            List<string> targetlist = new List<string>();
            foreach (var ip in IPAddressRange.Parse(TargetString))
                targetlist.Add(ip.ToString());
            return targetlist;
        }
        else
            return new List<string> { TargetString };
    }

    //Function to Live Hosts via ICMP. Will update the target array so we don't continuously attempt to connect to down systems
    static List<string> LiveTargetDetection(List<string> Targets)
    {
        foreach (string Target in Targets.ToArray())
        {
            if (!PingHost(Target))
                Targets.Remove(Target);
        }
        return Targets;
    }

    //Start Function
    public static void Run(string[] args)
    {
        try
        {
            Console.WriteLine("-------------------------------START--------------------------------");
            Parser.Default.ParseArguments<Options>(args).WithParsed<Options>(o =>
            {
                List<string> Targets = TargetParsing(o.TargetString);
                if (Targets == null)
                    Console.WriteLine("Error: Targets is null");

                if (o.Verbose && !o.SkipDetection)
                    Console.WriteLine("Pinging target list to identify live systems");

                if (!o.SkipDetection)
                    Targets = LiveTargetDetection(Targets);

                if (o.Verbose)
                {
                    Console.WriteLine("-----------------------------PARAMETERS-----------------------------");
                    Console.WriteLine("Targets: {0}", string.Join(",", Targets));
                    Console.WriteLine("Poll Amount: {0}", o.PollAmount);
                    Console.WriteLine("Will Poll for {0} minutes (30 second poll frequency * poll amount)", o.PollAmount * 0.5);
                    Console.WriteLine("--------------------------------------------------------------------");
                }

                Console.WriteLine("Creating profiles.csv with target information");
                foreach (string Target in Targets.ToArray())
                {
                    if (o.Verbose)
                        Console.WriteLine("Populating profiles.csv for {0}", Target);

                    ManagementScope msTargetScope = ScopeConnect("\\\\" + Target + "\\root\\Cimv2", 1);
                    if (msTargetScope != null)
                    {
                        CsvFileWriter writer = new CsvFileWriter("profiles.csv", true);
                        GenerateTargetCSV(msTargetScope, writer);
                        writer.Dispose();                        
                    }
                    else
                    {
                        Console.WriteLine("Error connecting to {0}", Target);
                        if (o.AutoRemove)
                        {
                            if (o.Verbose)
                                Console.WriteLine("Removing {0} from target list ", Target);
                            Targets.Remove(Target);
                        }
                    }                   
                }

                Console.WriteLine("Creating networkconnections.csv with connection information");
                int iCounter = 0;
                while (iCounter < o.PollAmount && Targets.Count != 0)
                {
                    foreach (string Target in Targets.ToArray())
                    {
                        if (o.Verbose)
                            Console.WriteLine("Populating networkconnections.csv for {0}", Target);
                        ManagementScope msConnectionScope = ScopeConnect("\\\\" + Target + "\\root\\StandardCimv2", 1);
                        if (msConnectionScope != null)
                        {
                            CsvFileWriter writer = new CsvFileWriter("networkconnections.csv", true);
                            CollectNetworkConnections(msConnectionScope, writer);
                            writer.Dispose();
                        }
                        else
                        {
                            Console.WriteLine("Error connecting to {0}", Target);
                            if (o.AutoRemove)
                            {
                                if (o.Verbose)
                                    Console.WriteLine("Removing {0} from target list ", Target);
                                Targets.Remove(Target);
                            }
                        }
                    }
                    Thread.Sleep(30000);
                    iCounter++;
                }
            });
            return;
        }
        catch (Exception e)
        {
            Console.WriteLine("{0} Expection Caught.", e);
        }        
    }
}

public class CsvRow : List<string>
{
    public string LineText { get; set; }
}
public class CsvFileWriter : StreamWriter
{
    public CsvFileWriter(Stream stream)
        : base(stream)
    {
    }

    public CsvFileWriter(string filename)
        : base(filename)
    {
    }

    public CsvFileWriter(string path, bool append)
      : base(path, append)
    {
    }

    /// <summary>
    /// Writes a single row to a CSV file.
    /// </summary>
    /// <param name="row">The row to be written</param>
    public void WriteRow(CsvRow row)
    {
        StringBuilder builder = new StringBuilder();
        bool firstColumn = true;
        foreach (string value in row)
        {
            // Add separator if this isn't the first value
            if (!firstColumn)
                builder.Append(',');
            // Implement special handling for values that contain comma or quote
            // Enclose in quotes and double up any double quotes
            if (value.IndexOfAny(new char[] { '"', ',' }) != -1)
                builder.AppendFormat("\"{0}\"", value.Replace("\"", "\"\""));
            else
                builder.Append(value);
            firstColumn = false;
        }
        row.LineText = builder.ToString();
        WriteLine(row.LineText);
    }
}

namespace ThunderQuery
{   
    class Program
    {
        static void Main(string[] args)
        {            
            ThunderQ.Run(args);
            return;
        }
        
    }
}
