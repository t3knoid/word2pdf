# word2pdf
word2pdf is a Windows command-line application that converts Word document to PDF. It uses Microsoft Office interop to automate saving a given word file into a PDF document.

# Requirement
Microsoft Word must be installed on the system where word2pdf is executed.

# Special Considerations when executing from a service under 64-bit Windows
When using this utility under a service such as Jenkins under 64-bit Windows, the system must be configured properly, otherwise documents cannot be opened. Typically the application will display the message, "Object reference not set to an instance of an object." After extensive searching on the internet, I found the following thread the provided a solution:

https://social.msdn.microsoft.com/Forums/en-US/0f5448a7-72ed-4f16-8b87-922b71892e07/word-2007-documentsopen-returns-null-in-aspnet?forum=architecturegeneral

1. Launch C:\Windows\System32\comexp.msc from the system with the issue.
2. Traverse > Component Services > Computers > My Computer > DCOM Config > Microsoft Word 97-2003 Document.
3. Right-click the Microsoft Word 97-2003 Document entry and select Properties.
4. Click on the Identity tab.
5. Select "The interactive user."

