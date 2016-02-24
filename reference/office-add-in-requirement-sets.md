
# Office add-in requirement sets
List of all Office add-in requirement sets, and methods that do not belong to a requirement set in Office.js.

Requirement sets are named groups of API members. Office add-ins use requirement sets specified in the manifest or using a runtime check to determine if an Office host supports APIs needed by the add-in. For more information, see [Specify Office hosts and API requirements](http://msdn.microsoft.com/library/6b6702f2-b0a5-46ab-a356-8dda897ca8ae%28Office.15%29.aspx).


## Requirement sets
<a name="SpecifyRequirementSets_sets"> </a>

The following table lists the names of requirement sets, the methods in each set, the Office host applications that support that requirement set, and the version number of the API.



|**Set name**|**Version**|**Methods in set**|**Office host**|
|:-----|:-----|:-----|:-----|
|ExcelApi|1.1|All elements in the Excel namespace.|Excel 2016Excel Online|
|WordApi|1.2|All elements in the Word namespace. The following methods were added to this version of WordApi:Body.select(selectionMode)Body.insertInlinePictureFromBase64(base64EncodedImage, insertLocation)contentControl.select(selectionMode)contentControl.insertInlinePictureFromBase64(base64EncodedImage, insertLocation)inlinePicture.paragraphinlinePicture.deleteinlinePicture.insertBreak(breakType, insertLocation)inlinePicture.insertFileFromBase64(base64file, insertLocation)inlinePicture.insertHtml(html, insertLocation)inlinePicture.insertInlinePictureFromBase64(base64file, insertLocation)inlinePicture.insertOoxml(ooxml, insertLocation)inlinePicture.insertParagraph(paragraphText, insertLocation)inlinePicture.insertText(text, insertLocation)inlinePicture.select(selectionMode)paragraph.select(selectionMode)range.inlinePicturesrange.select(selectionMode)range.insertInlinePictureFomBase64(base64EcodedImage, insertLocation)|Word 2016|
|WordApi|1.1|All elements in the Word namespace except API members that were added to WordApi 1.2 and later, which are listed above.|Word 2016|
|ActiveView|1.1|Document.getActiveViewAsync|PowerPointPowerPoint Online|
|BindingEvents|1.1|Binding.addHanderAsyncBinding.removeHanderAsync|Access web appsExcelExcel OnlineWord|
|CompressedFile|1.1|Supports output to Office Open XML (OOXML) format as a byte array (Office.FileType.Compressed) when using the Document.getFileAsync method.|PowerPointWord|
|CustomXmlParts|1.1|CustomXmlNode.getNodesAsyncCustomXmlNode.getNodeValueAsyncCustomXmlNode.getXmlAsyncCustomXmlNode.setNodeValueAsyncCustomXmlNode.setXmlAsyncCustomXmlPart.addHandlerAsyncCustomXmlPart.deleteAsyncCustomXmlPart.getNodesAsyncCustomXmlPart.getXmlAsyncCustomXmlPart.removeHandlerAsyncCustomXmlParts.addAsyncCustomXmlParts.getByIdAsyncCustomXmlParts.getByNamespaceAsyncCustomXmlPrefixMappings.addNamespaceAsyncCustomXmlPrefixMappings.getNamespaceAsyncCustomXmlPrefixMappings.getPrefixAsync|Word|
|DocumentEvents|1.1|Document.addHandlerAsyncDocument.removeHandlerAsync|ExcelExcel OnlinePowerPointWordWord Online|
|File|1.1|Document.getFileAsyncFile.closeAsyncFile.getSliceAsync|PowerPointWordWord Online|
|HtmlCoercion|1.1|Supports coercion to HTML (Office.CoercionType.Html) when reading and writing data using the Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync, or Binding.setDataAsync methods.|Word|
|ImageCoercion|1.1|Supports conversion to an image (Office.CoercionType.Image) when writing data using the Document.setSelectedDataAsync method.|WordWord Online|
|Mailbox|1.1|All API members supported by Outlook add-ins (those members accessed from  `Office.context` and `Office.context.mailbox` in your add-in's code).|OutlookOutlook Web AppOWA for Devices|
|MatrixBindings|1.1|Bindings.addFromNamedItemAsyncBindings.addFromSelectionAsyncBindings.getAllAsyncBindings.getByIdAsyncBindings.releaseByIdAsyncMatrixBinding.getDataAsyncMatrixBinding.setDataAsync|ExcelExcel OnlineWord|
|MatrixCoercion|1.1|Supports coercion to the "matrix" (array of arrays) data structure (Office.CoercionType.Matrix) when reading and writing data using the Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync, or Binding.setDataAsync methods.|ExcelExcel OnlineWord|
|OoxmlCoercion|1.1|Supports coercion to Open Office XML (OOXML) format (Office.CoercionType.Ooxml) when reading and writing data using the Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync, or Binding.setDataAsync methods.|Word|
|PartialTableBindings|1.1||Access web apps|
|PdfFile|1.1|Supports output to PDF format (Office.FileType.Pdf) when using the Document.getFileAsync method.|PowerPointWord|
|Selection|1.1|Document.getSelectedDataAsyncDocument.setSelectedDataAsync|ExcelExcel OnlinePowerPointProjectWord|
|Settings|1.1|Settings.getSettings.removeSettings.saveAsyncSettings.set|Access web appsExcelExcel OnlinePowerPointPowerPoint OnlineWordWord Online|
|TableBindings|1.1|Bindings.addFromNamedItemAsyncBindings.addFromSelectionAsyncBindings.getAllAsyncBindings.getByIdAsyncBindings.releaseByIdAsyncTableBinding.addColumnsAsyncTableBinding.addRowsAsyncTableBinding.deleteAllDataValuesAsyncTableBinding.getDataAsyncTableBinding.setDataAsync|Access web appsExcelExcel OnlineWord|
|TableCoercion|1.1|Supports coercion to the "table" data structure (Office.CoercionType.Table) when reading and writing data using the Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync, or Binding.setDataAsync methods.|Access web appsExcelExcel OnlineWord|
|TextBindings|1.1|Bindings.addFromNamedItemAsyncBindings.addFromSelectionAsyncBindings.getAllAsyncBindings.getByIdAsyncBindings.releaseByIdAsyncTextBinding.getDataAsyncTextBinding.setDataAsync|ExcelExcel OnlineWord|
|TextCoercion|1.1|Supports coercion to text format (Office.CoercionType.Text) when reading and writing data using the Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync, or Binding.setDataAsync methods.|ExcelExcel OnlinePowerPointProjectWordWord Online|
|TextFile|1.1|Supports output to text format (Office.FileType.Text) when using the Document.getFileAsync method.|Word|

## Methods that aren't part of a requirement set
<a name="SpecifyRequirementSets_methods"> </a>

The following methods in the JavaScript API for Office aren't part of a requirement set. If your add-in requires any of these methods, use the  **Methods** and **Method** elements in the add-in's manifest to declare that they are required, or perform the runtime check using an if statement. For more information, see[Specify Office hosts and API requirements](http://msdn.microsoft.com/library/6b6702f2-b0a5-46ab-a356-8dda897ca8ae%28Office.15%29.aspx).



|**Method name**|**Office host support**|
|:-----|:-----|
|Bindings.addFromPromptAsync|Access web apps, Excel and Excel Online|
|Document.getFilePropertiesAsync|Excel, Excel Online, Word, and PowerPoint|
|Document.getProjectFieldAsync|Project Standard 2013 and Project Professional 2013|
|Document.getResourceFieldAsync|Project Standard 2013 and Project Professional 2013|
|Document.getSelectedResourceAsync|Project Standard 2013 and Project Professional 2013|
|Document.getSelectedTaskAsync|Project Standard 2013 and Project Professional 2013|
|Document.getSelectedViewAsync|PowerPoint and PowerPoint Online|
|Document.getTaskAsync|Project Standard 2013 and Project Professional 2013|
|Document.getTaskFieldAsync|Project Standard 2013 and Project Professional 2013|
|Document.goToByIdAsync|Excel, Excel Online, Word, and PowerPoint|
|Settings.addHandlerAsync|Access web apps, Excel, Excel Online, Word, and PowerPoint|
|Settings.refreshAsync|Access web apps, Excel, Excel Online, Word, PowerPoint, and PowerPoint Online|
|Settings.removeHandlerAsync|Access web apps, Excel, Excel Online, Word, and PowerPoint|
|TableBinding.clearFormatsAsync|Excel, Excel Online|
|TableBinding.setFormatsAsync|Excel, Excel Online|
|TableBinding.setTableOptionsAsync|Excel, Excel Online|

## Additional resources
<a name="bk_addresources"> </a>


- [Specify Office hosts and API requirements](http://msdn.microsoft.com/library/6b6702f2-b0a5-46ab-a356-8dda897ca8ae%28Office.15%29.aspx)
    
