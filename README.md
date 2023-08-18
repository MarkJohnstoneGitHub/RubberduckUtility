# RubberduckUtility
[Rubberduck](https://rubberduckvba.com/) utility to export all components according to [@Folder annotation](https://github.com/rubberduck-vba/Rubberduck/wiki/Using-@Folder-Annotations). 

Required references: VBIDE (Microsoft Visual Basic for Applications Extensibility 5.3)

Usage: eg. ```RubberduckUtility.ExportAllComponents "C:\VBA\Output\ "```


Added error handling for: 

Invalid working directory.  
  - Error is raised.  The working directory must already exist.

Rubberduck @Folder annotations that contain invalid folder characters.  
  - Components containing  invalid folder annotations are exported to the working directory.
  eg. '@Folder "<Rubberduck Utilities>"
  - A warning message is displayed in the immediate window.
  - Invalid rubberduck folder annotation, <Rubberduck Utilities> RubberDuckExport.bas exported to working directory C:\VBA\Output\

