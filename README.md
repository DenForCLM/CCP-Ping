# CCP-Ping
Ping all network interfaces from the CCP (MOSAIQ) PC

1. Placement

This script should be placed on the CCP (MOSAIQ) PC.
A common choice is to store it in a folder like C:\Temp or C:\Scripts (you can pick any convenient folder).

2. Launching

Create a desktop shortcut pointing to this script.
Double-click the shortcut to run it.

3. Function

The script pings a predefined list of servers/network devices.
It saves the ping results into a text file (named with the current date and time) in the same folder where the script is located.
Once finished, the script automatically opens that text file in Notepad.

Output file:
IP                Status  Device            Date/Time           
192.168.30.1      up      NSS               1/15/2025 6:36:49 PM
192.168.30.2      up      Integrity_VM      1/15/2025 6:36:49 PM
192.168.30.3      up      Mosaiq            1/15/2025 6:36:49 PM
192.168.30.4      up      XVI               1/15/2025 6:36:49 PM
192.168.30.5      up      iViewGT           1/15/2025 6:36:49 PM
192.168.30.7      up      iGuide            1/15/2025 6:36:49 PM
192.168.30.16     up      UPS               1/15/2025 6:36:49 PM
192.168.30.17     up      NAS               1/15/2025 6:36:49 PM
192.168.30.150    up      TRM_Computer      1/15/2025 6:36:49 PM
192.168.30.200    up      CCP-Management    1/15/2025 6:36:49 PM
192.168.81.2      up      IntelliMax_VM     1/15/2025 6:36:49 PM
192.168.240.244   up      Netgear_switch    1/15/2025 6:36:49 PM
192.168.240.247   up      VM_access         1/15/2025 6:36:49 PM
192.168.240.250   up      NRT_server        1/15/2025 6:36:49 PM

local IP addresses list:
192.168.30.200
192.168.30.3

