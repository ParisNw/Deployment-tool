# PC Deployment Tool

This tool is used to quickly setup AD attributes for multiple machines at the same time. It is a Windows Form, written in Powershell and uses Active Directory cmdlets search for or configure computers in AD.

## Installation

* At least .NET 3.5

* Powershell, which is installed by default on all Win 7 (SP1) or later machines.

* [RSAT](https://support.microsoft.com/en-us/help/2693643/remote-server-administration-tools-rsat-for-windows-operating-systems), as the buttons run cmdlets available with the [Active Directory module for Windows Powershell](https://docs.microsoft.com/en-us/powershell/module/addsadministration/?view=win10-ps).

## Usage

1. Select a radial to filter the search results. Then use the textbox and click the magnifying glass to search.

2. Your results will populate in this box. From here you can select AD objects to queue. Hold shift to select multiple objects and then click the green **"+"** to send it to the corresponding queue

3. Next, clicking the "Move to OU" or "Add Groups" buttons will add all of the queued devices to the OU or groups, respectively.
