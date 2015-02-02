# LOSessionTarget
Libreoffice Writer Session Word Count with Style Macro

Set a word count target.
Choose what paragraph styles to include in the word count.  (Default Style is actually named Standard).
Useful when writing from bullet point outline that you delete once the topic has been covered.  Those deleted words will not be included in the word count if those notes are in a paragraph style not listed in the macro dialog box.  Word Count resets every time macro is run.

1.  Copy wordCountbyStyle.py into Users\$USER$\AppData\Roaming\LibreOffice\4\user\Scripts\python (or appropriate folder for LO version and OS).
2.  To run Tools -> Macros -> Run Macro; My Macros (Library), WordCountbyStyle, wordCountbyStyle (Macro name).
3.  Word Count resets every time macro is run.
4.  Word Count Target is stored in user field found under Insert -> Fields -> Other; Variables tab.
