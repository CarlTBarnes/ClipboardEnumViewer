# ClipboardEnumViewer - Enumerate Clipboard Formats in Clarion

Simple program to call EnumClipboardFormats() and show all the formats. Can also view the clipboard data using the Clarion Clipboard() function. Calling the Windows API will be needed for many Formats.

Below is a screen capture of an Excel copy that shows many formats.

![screen cap](readme1.png)

Excel Example:

Format# | Format Name
--------|------------
49161 | DataObject
14 | CF_ENHMETAFILE
3 | CF_METAFILEPICT
2 | CF_BITMAP
50009 | Biff12
49986 | Biff8
49988 | Biff5
4 | CF_SYLK
5 | CF_DIF
50007 | XML Spreadsheet
49381 | HTML Format
13 | CF_UNICODETEXT
1 | CF_TEXT
49989 | Csv
50006 | Hyperlink
49308 | Rich Text Format
49163 | Embed Source
49156 | Native
49155 | OwnerLink
49166 | Object Descriptor
49165 | Link Source
49167 | Link Source Descriptor
49985 | Link
129 | CF_DSPTEXT
49154 | ObjectLink
49171 | Ole Private Data
16 | CF_LOCALE
7 | CF_OEMTEXT
8 | CF_DIB
17 | CF_MAX
