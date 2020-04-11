[![License](https://img.shields.io/github/license/rkapl123/CmdLogAddin.svg)](https://github.com/rkapl123/CmdLogAddin/blob/master/LICENSE)

# CmdLogAddin

Excel Addin that allows you to parse Excel's Cmdline and start any Macro that is contained either inside the started Workbook, a startup loaded Addin or outside.  

Additionally, a logging possibility is provided by retrieving a logger object in VBA (set log = CreateObject("LogAddin.Logger")) and using this to
provide logging messages using 5 levels:  

- log.Fatal (like log.Error but quits Excel)
- log.Error (also can send Mails, if desired)
- log.Warn
- log.Info
- log.Debug

Details see https://rkapl123.github.io/CmdLogAddin/
