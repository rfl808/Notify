# Notify
Open your workbook in Excel.
Press Alt + F11 to open Visual Basic Editor (VBE).
Right-click on your workbook name in the "Project-VBAProject" pane (at the top left corner of the editor window) and select Insert -> Module from the context menu.
Copy the VBA notify code and paste it to the right pane of the VBA editor ("Module1" window).

Syntax

notify (titleMessage, [infoMessage], [flagOfMessage])

'Flags for the balloon message..
'None = 0
'Information = 1
'Exclamation = 2
'Critical = 3