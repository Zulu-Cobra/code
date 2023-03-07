' Set the network path to the file share
Dim networkPath
networkPath = "\\143.42.198.172\share"

' Set the username and password for the file share
Dim username
Dim password
username = "username"
password = "password"

' Create a network object
Dim network
Set network = WScript.CreateObject("WScript.Network")

' Map the network drive
network.MapNetworkDrive "", networkPath, False, username, password

' Display a message box to indicate success
MsgBox "Connected to file share: " & networkPath

' Do some work on the network drive here...

' Disconnect the network drive
network.RemoveNetworkDrive networkPath, True, True

' Display a message box to indicate success
MsgBox "Disconnected from file share: " & networkPath
