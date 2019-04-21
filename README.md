ThunderQuery
------------
A C# application which enumerates network connections on targets, a netstat via WMI.

References
------------
System.Management

Command Line Parser - https://github.com/commandlineparser/commandline
PM> Install-Package CommandLineParser

IP Address Range - https://github.com/jsakamoto/ipaddressrange
PM> Install-Package IPAddressRange


Build Notes
------------
Created using Visual Studio 2017 and .NET Framework 4.7. If I recall correctly, Command Line Parser works .NET 4.0 so any framework 4.0+ should work fine. 
Just a FYI, while test building .NET 4.0 I had to tinker with the versions of CommandLineParser (2.3.0) and IPAddressRange (1.6.0).

Additionally, I used ilmerge.exe to embed the assemblies, as follows:
"c:\Program Files (x86)\Microsoft\ILMerge\ILMerge.exe" /targetplatform:v4 ThunderQuery.exe CommandLine.dll IPAddressRange.dll /out:ThunderQueryFinal.exe

Usage
------------
-t, --targets            (Default: 127.0.0.1) Targets to poll. Will accept comma separated, hyphenated, or single targets.
						  For example, 192.168.1.50 ... 192.168.1.50-192.168.1.100 ... 192.168.1.50,192.168.1.51,192.168.1.52

-p, --pollamount         (Default: 10) Amount of times tool will poll targets every 30 seconds. Default 10 times - 10 * 30 seconds = 5 minutes.

-s, --skipdetection      (Default: false) Tool will ping systems before attempting to enumerate via WMI. Use this switch to skip live detection.

-a, --autoremove         (Default: true) Will automatically remove systems which fail WMI query.

-v, --verbose            (Default: false) Be verbose

--help                   Display this help screen.

--version                Display version information.

What It Does
------------
Will remotely enumerate targets via WMI to develop a system profile per target and store it in profiles.csv, as follows:
	-Hostname
	-Domain
	-IP Address
	-LoggedOn User
	-Domain Role

After profiles.csv is populated the tool will remotely poll targets every 30 seconds via WMI to collect established TCP connections. Remote and Local port numbers above 49152 are filtered out and stored in network connections.csv.
WMI Query - SELECT LocalAddress,LocalPort,RemoteAddress,RemotePort FROM MSFT_NetTCPConnection WHERE state = 5 AND RemoteAddress != '127.0.0.1' AND RemoteAddress != '::1' AND (RemotePort < 49152 OR LocalPort < 49152)

Result
------------
Will create two CSV files: profiles.csv and networkconnections.csv

profiles.csv - Hostname,Domain,LoggedOn User,DomainRole,IPAddress
Example - BLANPC-0004,blan.local,BLAN\Jack,DomainMember_Wks,10.10.112.55

networkconnections.csv - SourceIP,SourcePort,DestinationIP,DestinationPort
Example - 10.10.112.55,52847,10.10.112.50,135

Neo4j
------------
Importing into Neo4j (See https://ijustwannared.team for further details):

LOAD CSV FROM "file:///profiles.csv" AS row1 CREATE(:DNSHostName:IPAddress{dnshostname:row1[0],domain:row1[1], username: row1[2],domainrole: row1[3],ipaddress:row1[4]})

LOAD CSV FROM "file:///networkconnections.csv" AS row2 MERGE(:IPAddress{ipaddress:row2[2]})

LOAD CSV FROM "file:///networkconnections.csv" AS row3 MATCH (a:IPAddress),(b:IPAddress) WHERE a.ipaddress = row3[0] AND b.ipaddress = row3[2] AND tointeger(row3[3]) < 49152 MERGE (a)-[r:NetConn {DestPort: tointeger(row3[3])}]->(b) return r

LOAD CSV FROM "file:///networkconnections.csv" AS row4 MATCH (a:IPAddress),(b:IPAddress) WHERE a.ipaddress = row4[0] AND b.ipaddress = row4[2] AND tointeger(row4[1]) < 49152 MERGE (b)-[r:NetConn {DestPort: tointeger(row4[1])}]->(a) return r
