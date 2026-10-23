; XD3 PDF Editor — installer additions (electron-builder "include").
; Lists the app in "Open with" and in Windows' Default apps for PDF and Word files, for the user who installs it.
; The installer never makes itself the default PDF app: the app offers that and Windows asks the user.

!macro xd3Type ID DESC
  WriteRegStr HKCU "Software\Classes\${ID}" "" "${DESC}"
  WriteRegStr HKCU "Software\Classes\${ID}" "FriendlyTypeName" "${DESC}"
  WriteRegStr HKCU "Software\Classes\${ID}\DefaultIcon" "" '"$INSTDIR\${APP_EXECUTABLE_FILENAME}",0'
  WriteRegStr HKCU "Software\Classes\${ID}\Application" "ApplicationName" "XD3 PDF Editor"
  WriteRegStr HKCU "Software\Classes\${ID}\Application" "ApplicationCompany" "XDCybertech Pvt Ltd"
  WriteRegStr HKCU "Software\Classes\${ID}\shell" "" "open"
  WriteRegStr HKCU "Software\Classes\${ID}\shell\open" "FriendlyAppName" "XD3 PDF Editor"
  WriteRegStr HKCU "Software\Classes\${ID}\shell\open\command" "" '"$INSTDIR\${APP_EXECUTABLE_FILENAME}" "%1"'
!macroend

!macro customInstall
  !insertmacro xd3Type "XD3PDFEditor.pdf" "PDF Document"
  !insertmacro xd3Type "XD3PDFEditor.docx" "Word Document"
  WriteRegStr HKCU "Software\Classes\.pdf\OpenWithProgids" "XD3PDFEditor.pdf" ""
  WriteRegStr HKCU "Software\Classes\.docx\OpenWithProgids" "XD3PDFEditor.docx" ""
  WriteRegStr HKCU "Software\Classes\Applications\${APP_EXECUTABLE_FILENAME}" "FriendlyAppName" "XD3 PDF Editor"
  WriteRegStr HKCU "Software\Classes\Applications\${APP_EXECUTABLE_FILENAME}\shell\open\command" "" '"$INSTDIR\${APP_EXECUTABLE_FILENAME}" "%1"'
  WriteRegStr HKCU "Software\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".pdf" ""
  WriteRegStr HKCU "Software\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".docx" ""
  WriteRegStr HKCU "Software\XD3Labs\XD3 PDF Editor\Capabilities" "ApplicationName" "XD3 PDF Editor"
  WriteRegStr HKCU "Software\XD3Labs\XD3 PDF Editor\Capabilities" "ApplicationDescription" "Free offline PDF viewer and editor, with a Word editor"
  WriteRegStr HKCU "Software\XD3Labs\XD3 PDF Editor\Capabilities" "ApplicationIcon" '"$INSTDIR\${APP_EXECUTABLE_FILENAME}",0'
  WriteRegStr HKCU "Software\XD3Labs\XD3 PDF Editor\Capabilities\FileAssociations" ".pdf" "XD3PDFEditor.pdf"
  WriteRegStr HKCU "Software\XD3Labs\XD3 PDF Editor\Capabilities\FileAssociations" ".docx" "XD3PDFEditor.docx"
  WriteRegStr HKCU "Software\RegisteredApplications" "XD3 PDF Editor" "Software\XD3Labs\XD3 PDF Editor\Capabilities"
  System::Call 'shell32::SHChangeNotify(i 0x08000000, i 0, p 0, p 0)'
!macroend

!macro customUnInstall
  DeleteRegValue HKCU "Software\RegisteredApplications" "XD3 PDF Editor"
  DeleteRegKey HKCU "Software\XD3Labs\XD3 PDF Editor"
  DeleteRegKey /ifempty HKCU "Software\XD3Labs"
  DeleteRegValue HKCU "Software\Classes\.pdf\OpenWithProgids" "XD3PDFEditor.pdf"
  DeleteRegValue HKCU "Software\Classes\.docx\OpenWithProgids" "XD3PDFEditor.docx"
  ReadRegStr $0 HKCU "Software\Classes\.pdf" ""
  StrCmp $0 "XD3PDFEditor.pdf" 0 +2
    DeleteRegValue HKCU "Software\Classes\.pdf" ""
  DeleteRegKey HKCU "Software\Classes\XD3PDFEditor.pdf"
  DeleteRegKey HKCU "Software\Classes\XD3PDFEditor.docx"
  DeleteRegKey HKCU "Software\Classes\Applications\${APP_EXECUTABLE_FILENAME}"
  System::Call 'shell32::SHChangeNotify(i 0x08000000, i 0, p 0, p 0)'
!macroend
