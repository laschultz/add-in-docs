
# LabsJS.Labs.Components
Provides a high-level overview of the Labs.JS Labs.Components JavaScript API.

 _**Applies to:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

APIs in the Labs.Components module represent the four default components that are presently available for the development of labs (Activity, Choice, Input, and Dynamic components).

## Labs.Components module

Following are Labs.Components types:


### Classes


|||
|:-----|:-----|
|[Labs.Components.ComponentAttempt](../powerpoint/office-mix/reference/labs.components.componentattempt.md)|Base class for attempts at components.|
|[Labs.Components.ActivityComponentAttempt](../powerpoint/office-mix/reference/labs.components.activitycomponentattempt.md)|Represents an attempt at completing an activity component.|
|[Labs.Components.ActivityComponentInstance](../powerpoint/office-mix/reference/labs.components.activitycomponentinstance.md)|Represents the current instance of an activity component.|
|[Labs.Components.ChoiceComponentAnswer](../powerpoint/office-mix/reference/labs.components.choicecomponentanswer.md)|The answer to a problem presented in a choice component.|
|[Labs.Components.ChoiceComponentAttempt](../powerpoint/office-mix/reference/labs.components.choicecomponentattempt.md)|Represents an attempt at a choice component.|
|[Labs.Components.ChoiceComponentInstance](../powerpoint/office-mix/reference/labs.components.choicecomponentinstance.md)|Represents an instance of a choice component.|
|[Labs.Components.ChoiceComponentResult](../powerpoint/office-mix/reference/labs.components.choicecomponentresult.md)|The result of a choice component submission.|
|[Labs.Components.ChoiceComponentSubmission](../powerpoint/office-mix/reference/labs.components.choicecomponentsubmission.md)|Represents the submission associated with a choice component.|
|[Labs.Components.DynamicComponentInstance](../powerpoint/office-mix/reference/labs.components.dynamiccomponentinstance.md)|Represents an instance of a dynamic component.|
|[Labs.Components.InputComponentAnswer](../powerpoint/office-mix/reference/labs.components.inputcomponentanswer.md)|Represents the answer to an input component problem.|
|[Labs.Components.InputComponentAttempt](../powerpoint/office-mix/reference/labs.components.inputcomponentattempt.md)|Represents an attempt at interacting with an input component.|
|[Labs.Components.InputComponentInstance](../powerpoint/office-mix/reference/labs.components.inputcomponentinstance.md)|Represents an instance of an input component.|
|[Labs.Components.InputComponentResult](../powerpoint/office-mix/reference/labs.components.inputcomponentresult.md)|The result of an input component submission.|
|[Labs.Components.InputComponentSubmission](../powerpoint/office-mix/reference/labs.components.inputcomponentsubmission.md)|Represents a submission to an input component.|

### Interfaces


|||
|:-----|:-----|
|[Labs.Components.IActivityComponent](../powerpoint/office-mix/reference/labs.components.iactivitycomponent.md)|Represents an activity component. Extends [Labs.Core.IComponent](../powerpoint/office-mix/reference/labs.core.icomponent.md).|
|[Labs.Components.IActivityComponentInstance](../powerpoint/office-mix/reference/labs.components.iactivitycomponentinstance.md)|Represents a specific instance of an activity component. Extends [Labs.Core.IComponentInstance](../powerpoint/office-mix/reference/labs.core.icomponentinstance.md).|
|[Labs.Components.IChoice](../powerpoint/office-mix/reference/labs.components.ichoice.md)|An available choice for a given problem.|
|[Labs.Components.IChoiceComponent](../powerpoint/office-mix/reference/labs.components.ichoicecomponent.md)|Enables interactions with a choice component.|
|[Labs.Components.IChoiceComponentInstance](../powerpoint/office-mix/reference/labs.components.ichoicecomponentinstance.md)|An instance of a choice component.|
|[Labs.Components.IDynamicComponent](../powerpoint/office-mix/reference/labs.components.idynamiccomponent.md)|Enables interaction with a dynamic component.|
|[Labs.Components.IDynamicComponentInstance](../powerpoint/office-mix/reference/labs.components.idynamiccomponentinstance.md)|An instance of a dynamic component.|
|[Labs.Components.IHint](../powerpoint/office-mix/reference/labs.components.ihint.md)|Hint for a lab problem.|
|[Labs.Components.IInputComponent](../powerpoint/office-mix/reference/labs.components.iinputcomponent.md)|Enables interacting with an input component.|
|[Labs.Components.IInputComponentInstance](../powerpoint/office-mix/reference/labs.components.iinputcomponentinstance.md)|An instance of an input component.|
