# RubberduckUtility
[Rubberduck](https://rubberduckvba.com/) utility to export all components according to Rubberduck [@Folder annotation](https://github.com/rubberduck-vba/Rubberduck/wiki/Using-@Folder-Annotations). 

Required references: VBIDE (Microsoft Visual Basic for Applications Extensibility 5.3)

Usage: eg. ```RubberduckUtility.ExportAllComponents "C:\VBA\Output\ "```

ExportAllComponents exports all components to the working directory provided for the active project, according to Rubberduck @Folder annotation. Sub folders are created according to the folder annotation and ***existing files are overwritten***.

Added error handling for: 

Invalid working directory.  
  - Error is raised.  The working directory must already exist.

Rubberduck @Folder annotations that contain invalid folder characters.  
  - Components containing @Folder annotations with invalid characters for folders are exported to the working directory.
  - eg. ``` '@Folder "<Rubberduck Utilities>" ```
  - A warning message is displayed in the immediate window.
  - Invalid rubberduck folder annotation, <Rubberduck Utilities> RubberDuckExport.bas exported to working directory C:\VBA\Output\

