# GetExchangeServerServicesReport.ps1
PowerShell script to generate a report of the Windows Services of an Exchange Server 2010/2013/2016 environment.

This week I had to figure out the status of our Windows services for Exchange on all servers in the Organisation.
I had no access to our Monitoring Systems and I was to lazy to login to all servers to check that manually and write the
report.
For this reason I created a small script which will do all that for me and which generates an HTML File.
Its also easy to forward that to my Boss.
The Report will be created with the Date of running and in the same Place from where you start the script.
