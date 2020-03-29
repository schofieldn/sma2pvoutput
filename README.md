# sma2pvoutput
Powershell script to download data from an SMA plant (inverter) and upload PV status information to PVOutput.org

This is a simple Powershell script to take the data from one or more SMA inverters and upload it to pvoutput.org. It requires the Sunny Explorer software from SMA to be installed on the Windows device on which the script is run in order to create the data exports.

The script can either take existing exports of status information for the specified date, or invoke Sunny Explorer to create an updated export. The contents of the export are then uploaded to pvoutput.org either individually or in batches.

It is designed to be run as a periodic scheduled task to perform status updates during daylight hours. The option then exists to schedule an end-of-day batch update of all the day's data to catch any statuses that may have been missed.

Best of luck.

Neil

The legal stuff:

The script is provided 'as is' without any warranty of any kind. SMA and Sunny Explorer are registered trademarks of SMA Solar Technologies AG.
