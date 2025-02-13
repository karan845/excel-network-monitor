# Excel Network Monitor

This Excel workbook (XLSM file) is designed to monitor the network status of multiple systems by pinging their IP addresses or hostnames. It leverages VBA (Visual Basic for Applications) to automate the process of sending ping requests and updating the worksheet in real time.

## Table of Contents

- [Introduction](#introduction)
- [Features](#features)
- [How It Works](#how-it-works)
- [Requirements](#requirements)
- [Usage Instructions](#usage-instructions)
- [Code Breakdown](#code-breakdown)
- [Customization](#customization)


## Introduction

The Excel Network Monitor is a simple yet effective tool that pings a list of systems and reports their online/offline status directly in an Excel sheet. It continuously checks each system until you choose to stop the process, making it useful for monitoring network devices or servers in a controlled environment.

## Features

- **Automated Network Monitoring:** Automatically pings a list of systems to determine their availability.
- **Real-Time Status Updates:** Updates the worksheet with the current status ("Online" or "Offline") for each system.
- **Visual Cues:** Uses color coding to indicate status changes—green for online and red for offline.
- **Stop Functionality:** Easily terminate the monitoring loop by setting a designated cell value.

## How It Works

1. **Input List:**  
   - Enter the IP addresses or hostnames of the systems you want to monitor in **Column B** (starting from row 2).

2. **Status Output:**  
   - The VBA code writes the current status ("Online" or "Offline") in **Column C** next to each IP/hostname.
   - It uses different font and background colors to highlight the system status.

3. **Monitoring Control:**  
   - **Cell F1** on the worksheet is used to control the monitoring process:
     - When set to `"TESTING"`, the code continues running.
     - To stop the monitoring process, change **F1** to `"STOP"`, or run the `stop_ping` macro.

## Requirements

- **Microsoft Excel:** Version supporting VBA and macros (e.g., Excel 2010 or later).
- **Windows Operating System:** The macro uses the Windows Script Host (`Wscript.Shell`) to execute the ping command.
- **Macros Enabled:** Ensure that Excel macros are enabled in your environment.

## Usage Instructions

1. **Setup:**
   - Open the XLSM file in Excel.
   - Populate **Column B** (starting from row 2) with the IP addresses or hostnames of the systems you wish to monitor.
   - Ensure that cell **F1** is not set to `"STOP"` (it should be blank or any other value) to allow the monitoring loop to start.

2. **Running the Macro:**
   - Press `Alt + F8` to open the Macro dialog.
   - Select `PingSystem` and click **Run**.
   - The macro will start pinging each system and update their status in **Column C**.

3. **Stopping the Monitoring:**
   - To stop the continuous loop, either:
     - Change the value in cell **F1** to `"STOP"`, or
     - Run the `stop_ping` macro via the Macro dialog.

4. **Viewing Results:**
   - Watch as each system's status is updated:
     - **Online:** Text will initially appear in black then change to green.
     - **Offline:** Text will appear in red, and the cell's background will change color.

## Code Breakdown

### `Ping` Function
```vb
Function Ping(strip)
    Dim objshell, boolcode
    Set objshell = CreateObject("Wscript.Shell")
    boolcode = objshell.Run("ping -n 1 -w 100 " & strip, 0, True)
    If boolcode = 0 Then
        Ping = True
    Else
        Ping = False
    End If
End Function
```
- **Purpose:** Sends a single ping request to the given IP/hostname.
- **Mechanism:** Uses the Windows Script Host to run the `ping` command with a 100 ms timeout.
- **Outcome:** Returns `True` if the ping is successful, otherwise `False`.

### `PingSystem` Subroutine
```vb
Sub PingSystem()
    Dim strip As String
    Do Until Sheet1.Range("F1").Value = "STOP"
        Sheet1.Range("F1").Value = "TESTING"
        For introw = 2 To ActiveSheet.Cells(65536, 2).End(xlUp).Row
            strip = ActiveSheet.Cells(introw, 2).Value
            If Ping(strip) = True Then
                ActiveSheet.Cells(introw, 3).Interior.ColorIndex = 0
                ActiveSheet.Cells(introw, 3).Font.Color = RGB(0, 0, 0)
                ActiveSheet.Cells(introw, 3).Value = "Online"
                Application.Wait (Now + TimeValue("0:00:01"))
                ActiveSheet.Cells(introw, 3).Font.Color = RGB(0, 200, 0)
            Else
                ActiveSheet.Cells(introw, 3).Interior.ColorIndex = 0
                ActiveSheet.Cells(introw, 3).Font.Color = RGB(200, 0, 0)
                ActiveSheet.Cells(introw, 3).Value = "Offline"
                Application.Wait (Now + TimeValue("0:00:01"))
                ActiveSheet.Cells(introw, 3).Interior.ColorIndex = 6
            End If
            If Sheet1.Range("F1").Value = "STOP" Then
                Exit For
            End If
        Next
    Loop
    Sheet1.Range("F1").Value = "IDLE"
End Sub
```
- **Purpose:** Continuously monitors the network status of systems listed in Column B.
- **Loop Control:** Uses cell **F1** as a flag to continue or break the monitoring loop.
- **Visual Feedback:** Updates the status and colors in Column C based on ping results.

### `stop_ping` Subroutine
```vb
Sub stop_ping()
    Sheet1.Range("F1").Value = "STOP"
End Sub
```
- **Purpose:** Provides an easy way to stop the `PingSystem` loop by setting cell **F1** to `"STOP"`.

## Customization

- **Timeout and Ping Count:** You can adjust the timeout value (`-w 100`) and the number of ping attempts (`-n 1`) within the `Ping` function to suit your network environment.
- **Worksheet References:** The code currently uses `Sheet1` and the active sheet. You may customize these references if your workbook has multiple sheets or different naming conventions.
- **Visual Cues:** Modify the RGB values or color index settings to change the visual representation of system statuses.

