
# LabsJS.Labs
Get an overview of APIs in the LabsJS.Labs module.

 _**Applies to:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

The LabsJS.Labs module contains the set of key JavaScript APIs that you can use to create the Office Add-ins (the labs). The APIs provide the entry point for lab development.

## LabsJS.Labs API module

The Labs module contains the following types:


### Variables


|||
|:-----|:-----|
|[Labs.DefaultHostBuilder](../powerpoint/office-mix/reference/labs.defaulthostbuilder.md)|Use this object to construct a default [Labs.Core.ILabHost](../powerpoint/office-mix/reference/labs.core.ilabhost.md) instance.|

### Functions


|||
|:-----|:-----|
|[Labs.Connect](../powerpoint/office-mix/reference/labs.connect.md)|Initializes a connection with the host.|
|[Labs.connect (overload)](../powerpoint/office-mix/reference/labs.connect-overload.md)|Initializes a connection with the host and provides input parameters.|
|[Labs.isConnected](../powerpoint/office-mix/reference/labs.isconnected.md)|Initializes a connection with the host.|
|[Labs.getConnectionInfo](../powerpoint/office-mix/reference/labs.getconnectioninfo.md)|Retrieves configuration information associated with a specified connection.|
|[Labs.disconnect](../powerpoint/office-mix/reference/labs.disconnect.md)|Disconnects the lab from the host and provides lab completion status.|
|[Labs.editLab](../powerpoint/office-mix/reference/labs.editlab.md)|Opens the specified lab for editing. You can specify the lab's configuration data while in edit mode. However, you cannot edit a lab while it is being taken (that is, the lab is running).|
|[Labs.takeLab](../powerpoint/office-mix/reference/labs.takelab.md)|Runs the specified lab and enables sending lab results to the server. Note that you cannot run a lab while it is being edited.|
|[Labs.on](../powerpoint/office-mix/reference/labs.on.md)|Adds a new handler for a specified event..|
|[Labs.off](../powerpoint/office-mix/reference/labs.off.md)|Removes an event handler for a specified event.|
|[Labs.getTimeline](../powerpoint/office-mix/reference/labs.gettimeline.md)|Retrieves a [Labs.Timeline](../powerpoint/office-mix/reference/labs.timeline.md) object instance that you can use to control the host player control.|
|[Labs.registerDeserializer](../powerpoint/office-mix/reference/labs.registerdeserializer.md)|Deserializes a specified JSON object into an object. Should be used by component authors only.|

### Classes


|||
|:-----|:-----|
|[Labs.ComponentInstanceBase](../powerpoint/office-mix/reference/labs.componentinstancebase.md)|Base class for the initialization of component instances.|
|[Labs.ComponentInstance](../powerpoint/office-mix/reference/labs.componentinstance.md)|Represents an instance of a component, which is an instantiation of a given component for a user at runtime. The object contains a translated view of the component for a specific run of a lab.|
|[Labs.Command](../powerpoint/office-mix/reference/labs.command.md)|General command used to pass messages between the client and host.|
|||
|[Labs.LabEditor](../powerpoint/office-mix/reference/labs.labeditor.md)|The  **LabEditor** object allows you to edit a given lab as well as get and set configuration data associated with the lab.|
|[Labs.LabInstance](../powerpoint/office-mix/reference/labs.labinstance.md)|An instance of a lab that is configured for the current user. Use this object to record and retrieve lab data for the user.|
|||
|||
|[Labs.Timeline](../powerpoint/office-mix/reference/labs.timeline.md)|Provides access to the labs.js timeline feature.|
|[Labs.ValueHolder](../powerpoint/office-mix/reference/labs.valueholder.md)|A container object that holds and tracks values for a specified lab. The value may be stored either locally or on the server.|

### Interfaces


|||
|:-----|:-----|
|[Labs.GetActionsCommandData](../powerpoint/office-mix/reference/labs.getactionscommanddata.md)|Allows you to retrieve data associated with a [LabsJS.Labs.Core.GetActions](../powerpoint/office-mix/reference/labsjs.labs.core.getactions.md) command.|
|[Labs.IMessageHandler](../powerpoint/office-mix/reference/labs.imessagehandler.md)|Interface that allows you to define event handlers.|
|[Labs.ITimelineNextMessage](../powerpoint/office-mix/reference/labs.itimelinenextmessage.md)|Provides means for interacting with the [Labs.Core.IMessage](https://msdn.microsoft.com/library/office/mt599680.aspx) object.|
|[Labs.SendMessageCommandData](../powerpoint/office-mix/reference/labs.sendmessagecommanddata.md)|Data associated with a [Labs.CommandType.TakeAction](https://msdn.microsoft.com/library/office/mt599680.aspx) command.|
|[Labs.TakeActionCommandData](../powerpoint/office-mix/reference/labs.takeactioncommanddata.md)|Data associated with a take action command.|

### Enumerations


|||
|:-----|:-----|
|[Labs.ConnectionState](../powerpoint/office-mix/reference/labs.connectionstate.md)|Enumerates the possible connection states of the lab to host.|
|[Labs.ProblemState](../powerpoint/office-mix/reference/labs.problemstate.md)||
