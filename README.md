# FroniusData
## View and backup historical data for Fronius inverters with Microsoft Excel

My free year of Fronius SolarWeb has expired.  Therefore, I created this tool to backup and view the energy production data from my 6kW solar tracker.

## Setup

This Excel workbook uses Visual Basic for Applications (VBA) to get realtime data and backup historical data for Fronius grid-tied inverters.  Within Excel you need to go into the Trust Center to enable VBA and macros, as follows:

**File -> Options -> Trust Center -> Trust Center Settings...**

*Macro Settings*
- Enable all macros (not recommended, potentially dangerous code can run)
- Trust access to the VBA project object model

You may also want to enable the Developer menu, to view and modify the VBA code.

**File -> Options -> Customize Ribbon**

*Main Tabs*
- Check the Developer menu


## Compatibility

I run Microsoft Excel 2019 on Windows 10.  There are some Windows API calls for various things, and I try to maintain compatibility with both 32-bit and 64-bit versions of Microsoft Office.

I have one Fronius inverter that feeds the electric grid for single-phase power (a.k.a.  North American split-phase power).  The inverter I use is this:  *Fronius Primo 6.0-1 208-240, Datalogger 3.23.6-1, SolarAPI v1.*


## Usage

To connect with the Fronius inverter on your local Ethernet or Wi-Fi network, you must enter either the "Inverter hostname" or "IP address" on the Excel worksheet.  If you do not know these pieces of information, the tool will search your local network for all Fronius inverters and show the results on the Devices tab.  Click **Update Realtime Data** to connect to the inverter.  If you did not initially type the "Inverter hostname" or "IP address", the tool will default the "IP address" to the first Fronius inverter found on your local network, and you will have to click the button a second time to connect.  If it works, you should see some information on the worksheet.

Great!  Now try viewing a Day chart for power generation.  Click the **Now** button next to the word "Day", to automatically pick today's date.  Then click **Show Chart**.

You may also click **Backup Historical Data** to copy all your energy production data from the inverter.  This may take some time, as the latest Fronius inverters can store up to 10 years of data with the 5 minute recording interval turned on.  My 1.5 years of data takes almost an hour to download.  However, once your data is downloaded, this tool never needs to ask the inverter again for this information. 

After backing up your historical data, an energy file is created and continuously maintained that tracks realtime power generation and cumulative energy production and consumption.  The file is called `energy.csv` and is located in the same folder as your historical data.  This is a flat file format that is easy to work with, compared to the structured JSON format returned by the Fronius inverter.

In the evening, the Fronius inverter goes into standby mode and this tool is designed for offline use. For this reason, historical data is stored on your computer. There is one JSON file for each day. These files are the exact JSON data returned by the Fronius inverter. The files are located here:

`%APPDATA%\FroniusData\inverter-{SERIAL-NUMBER}`

## Why?
I don't want to pay Fronius just to access my own data.  These solar trackers cost between $25K and $35K and it should come with free SolarWeb, just like my Garmin GPS has free lifetime maps!  The main difference between this tool and Fronius SolarWeb is that my historical data is stored locally on my computer, not in the cloud.  I chose to use Microsoft Excel because it was easy to get something up and running quickly while learning the Fronius SolarAPI.

