import win32com.client as win32
import re

outlook = win32.Dispatch("outlook.application")
olns = outlook.GetNamespace("MAPI")
inbox = olns.GetDefaultFolder(6).Folders.Item("Write-Off Responses")
# inbox = olns.GetDefaultFolder(5).Folders

filter = (
    "@SQL="
    + chr(34)
    + "urn:schemas:httpmail:subject"
    + chr(34)
    + " Like 'Write-Off' AND "
    + chr(34)
    + "urn:schemas:httpmail:hasattachment"
    + chr(34)
    + "=1"
)

# items = inbox.items.Restrict(filter)
items = inbox.items
for item in items:

    for attachment in item.Attachments:
        filename = attachment.Filename
        if str(filename).endswith(".xlsx"):
            print(filename)
            newname = re.findall("Write-Off \\d{3}", filename)
            newname = str(newname).replace('[', '').replace(']', '')
            attachment.SaveAsFile(
                f"C:\\users\\u6zn\\desktop\\writeoff\\Reviewed {newname}.xlsx"
            )
