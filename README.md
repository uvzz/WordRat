# WordRat

WordRAt runs in the background and leaks the content of open Word documents over HTTPS.

The script communicates with Word processes using the Marshal library's Word.Application COM object.

The script goes over all the open documents' contents using the Content.Text property, encodes the data with base64 (in case of special characters- to avoid syntax errors while sending over http), then uses paste.ee's API and exfiltrates the data with HTTP requests (by using the Msxml2.XMLHTTP com object). 

If Word is not running, the script waits for the process to get opened.
In case all Word processes are closed, the script runs again in a new process.  

I used an array to save leaked documents' names so the same document won't be sent more than once.
The script adds persistence by adding a registry value in the current user's "Run" registry key, and sets the value to a powershell encoded command, that runs the script at system startup. That way, no files are written to the disk and no high privileges are required.
