# Freeform Template Documentation

## Table of Contents
- [Forms](#forms)
- [Modules](#modules)
  * [TemplateActivation](#templateactivation)
  * [TemplateGlobalFunctions](#templateglobalfunctions)
  * [TemplateCopyLogic](#templatecopylogic)
  * [ConstParams](#constparams)
    + [Universal Constants](#universal-constants)
    + [Form Constants](#form-constants)
    + [General Params](#general-params)
    + [Parameter Array Declaration/Initialisation](#parameter-array-declarationinitialisation)
    + [Errors](#errors)
- [Editing Forms](#editing-forms)
  * [Adding Fields](#adding-fields)
    + [Changing Form Appearance](#changing-form-appearance)
      - [Field Settings](#field-settings)
      - [Naming Convention](#naming-convention)
      - [Form Code Updating](#updating-form-code)
        * [Populating Combo Boxes](#populating-combo-boxes)
        * [Linking Options with Children](#linking-options-with-children)
        * [Textbox Right Click](#textbox-right-click)
        * [Number-Only Text Input](#number-only-text-input)
    + [Updating Existing Elements](#updating-existing-elements)
      - [Cascading Name Values](#cascading-name-values)
      - [Updating TabIndex](#updating-tabindex)
    + [Updating Constant Parameters](#updating-constant-parameters)
    + [Updating Excel Sheet](#updating-excel-sheet)
  * [Removing Fields](#removing-fields)
- [Deploying for Production](#deploying-for-production)
  * [Rename File Copy](#rename-file-copy)
  * [Setup Exit on Close](#setup-exit-on-close)

## Forms
The Freeform Template application consists of five forms: `Default`, `Signal`, `Infrastructure`, `FO`, and `TO`. These forms have a small amount of code implemented within their respective objects, with the vast majority implemented in separate modules explained below.

## Modules
### TemplateActivation
The `TemplateActivation` module controls the now deprecated form on the `Form` sheet of the freeform excel file. It's functionality has been replaced by the pop up form controlled by the `ufrmMain` Form.

### TemplateGlobalFunctions
The `TemplateGlobalFunctions` module contains the functional code that is shared across all the forms.

| Function | Description |
| -------- | ----------- |
| `ClipBoard_SetData()` and `CopyTextToClipboard()` | The initial section declares the constants and functions required to enable the two functions used to copy data to the clip board, which follow directly. |
| `InitialiseForm()`  | Sets the `CURRENT_FORM_OBJ` and `CURRENT_FORM_ID` to refer to the object and form id of the most recent form to have been clicked. These global variables are then used throughout the code base for efficiency and modularity. |
| `GetInputValue()` | Provides a standard interface from extracting data from different field types. |
| `CheckRequiredInputs()` | Loops over the input fields of the active form, and checks whether those that are required have been filled. |
| `CheckSpelling()` | Loops over the input fields of the active form, and checks for spelling errors in spell check specified fields. |
| `CheckSpellingWords()` | Deprecated. |
| `ValidateTimeInputs()` | Evaluates whether the input in a time field is valid, according to the 24 hour `hhmm` clock format. This function does not throw an error for an empty input, as this would require all time inputs to be mandatory. |
| `ValidateComboBoxes()` | Ensures that the input in combo boxes is either empty, "YES", or "NO" (Case insensitive). An exception exists for the "Inward Crew/Stock Checked" combo box on the `TO` form, which is also allowed to have "N/A" as a valid input. |
| `SetActiveChildFields()` | When a combo box has it's value set or changed, this function de/activates all it's child fields according to the relationship specified in [the `initOptionFields` array](#parameter-array-declarationinitialisation). |
| `SetErrorText()` | Sets form error text field using standardised text for various error types. |
| `isInt()` | Evaluates whether an input value is numeric. |
| `isRadio()` | Determines whether an input is a combo box or not. |
| `extractTextAndCopy()` | Formats input text once error checking is complete. |
| `ShowTextMenu()`, `MakeMenu()`, and `TextMenu()` | Enable the use of a right click menu on forms, and copy and pasting content using this menu. |
   
### TemplateCopyLogic
The `TemplateCopyLogic` module defines the code that runs when the copy button is pressed on a form, in the `RunCopyLogic()` function. It handles the calls for error checking, input formatting, and copying to the clipboard by calling functions in the [`TemplateGlobalFunctions`](#templateglobalfunctions) module. It also feeds errors back to the form, updating label colours and error text for clarity. 

### ConstParams
The `ConstParams` module outlines the relationships and specifies the structure of each form. When changes are made to the application, almost all modifications to the code base occur in this file. As such, it is structured to be as easy manipulated and straightforward to understand as possible. It is broken down into four sections, which are explained in the rest of this chapter.

#### Universal Constants
The first section in this file specifies the constants that are not exclusive to any particular form. They are declared for convenience as well as for clarity and readability. The only cause for change would be when additional forms are created, at which point `NUM_FORMS` would need to be incremented. Additionally, a new form name constant would need to be created with the `FM_` prefix. This constant would then be used to configure the necessary parameters and specifications for the new form throughout the rest of the `ConstParams` module.

#### Form Constants
The form constants section is divided into two sub-sections. The first lists all the fields on that form, assigning the value of their order in the form (which corresponds with the number in the field's form ID) to a moniker that allows each field to easily be recognised when handled everywhere else in the code base.

The second sub-section assigns an additional moniker to those fields that are option fields, enabling a forms option fields to be iterated through in a convenient manner.

```vba
' Eg. Form Constants for TO Form
Public Const TO_INP_CTCHECKED = 1
Public Const TO_INP_SHOWS = 2
Public Const TO_INP_SSM_LOG = 3
...

Public Const TO_OPT_CTCHECKED = 1
Public Const TO_OPT_SSM_LOG = 2
...
```

> **Note:** Both of these sections must be updated when fields are added or removed to a form. Additionally, the values of the fields that follow must also be adjusted.

#### General Params
There are seven general parameters for each form, some or all of which will need to be updated when changes are made to the form they refer to.

| Field | Description |
| ------- | --------- |
| `NUM_OPTS` | The number of option fields (combo boxes) on the form |
| `NUM_FIELDS` | The number of fields of all types on the form |
| `NUM_COMMAS` | The number of line breaks in the form output |
| `FULL_RANGE` | The cells in the `field` sheet used to contain this form's input and titles |
| `CELL_RANGE` | The cells in the `field` sheet used to contain this form's input |
| `X_OFFSET` | Horizontal offset of this form's cells in the `field` sheet |
| `Y_OFFSET` | Vertical offset of this form's cells in the `field` sheet |

#### Parameter Array Declaration/Initialisation
Six arrays are used throughout the application to enforce certain parameter types and conditions that dictate form input. These arrays correspond to the following series of initialisation functions whose content must be configured to match their respective forms. The arrays are sized automatically according to the respective `NUM_FIELDS` value specified in the [General Params](#general-params) section, and each fields moniker can be used to specify whether it is active for that particular array's purpose:

```
' Set field CTCHECKED in FO as a required input
requiredInputs(FO_INP_CTCHECKED) = 1
```

* `initRequiredInputs()`

   Specifies which inputs will throw errors if not filled on a particular form.

* `initSpellCheckedInputs()`

   Specifies which inputs have active spell checking.  
   When a spell checking error is thrown, field is highlighted but incorrect text is not underlined.

* `initTimeTextInputs()`

   Specifies which inputs exclusively take time values formatted according to four digit 24h clock.

* `initLineBreakList()`

   Specifies which form fields to add a line break after in the input that gets sent to the clipboard.

* `initOptionInputs()`

   Specifies which fields are option fields (combo boxes), which require a 'yes' or 'no' input.

* `initOptionFields()`

   Some option fields have child questions which become mandatory to complete when a particular input is entered.
   Specifies which option fields have child fields.
   Allocates child fields to their parent field.
   Establishes whether child fields become activated on 'yes' or 'no' input.

#### Errors
Errors are accumulated in the `errorTypes` array. This array contains six integers, representing errors of different kinds, in order to facilitate the subsequent error messages that get displayed on the form. It is filled in the functions defined in the [`TemplateGlobalFunctions`](#templateglobalfunctions) module.

For errors defined by a custom set of conditions (such as particular combinations of fields on a form that must be filled in, or where at least one must be filled in), the `CheckSpecialErrors()` is used. This function has a space for custom error conditions to be defined for each form, assigned a type, and checked against the current submitted input.

Error messages are assigned in the [`TemplateCopyLogic()`](#templatecopylogic) function, terminating the function before any content is copied to the clipboard so that the user can rectify the problematic input. These messages are chosen according to the different error types that have been accumulated. These error types are defined in the [`ConstParams`](#constparams) module, and are described below. 

| Type | Description |
| ---- | ----------- |
| `ERROR_TIME` | Invalid value entered in time field |
| `ERROR_MULTI` | Particular combination of inputs needed |
| `ERROR_NOT_YES` | Field requires a positive response |
| `ERROR_REQUIRED` | Required field left empty |
| `ERROR_SPELLING` | Spelling error made in given field |
| `ERROR_INVALID` | Invalid option entered into combobox |

> **Note:** The `ERROR_MULTI` error type has no default associated message, and when used for a custom error, requires the accompanying error text to be manually created.

## Editing Forms

In order to successfully edit the application, a number of changes must be made to ensure that the various components coordinate as expected. The majority of changes are contained within the [`ConstParams`](#constparams) module, but in some cases changes in other section of the application may be required. This section outlines the two most likely situations in which changes will be made - adding and removing elements from forms.

### Adding Fields
#### Changing Form Appearance

The first stage of adding a field involves modifying the appearance of the form. Using the drag and drop editor, fields should be added in such a way that the visual layout of the form remains consistent and cohesive.

- Only text inputs and combo boxes should be used.

- Field dimensions should line up to the other elements on the form.

- Consistent spacing should be used.

- The same colour scheme should be used maintained.

- Mandatory fields should be indicated with an asterisk (*).

> *Note*: In some cases, it may be necessary to drastically alter the appearance of the form in order to accomodate needed changes. In this event, large numbers of fields should be organised into similarly sized groups using frames.

##### Field Settings

The drag and drop view has a window that allows field parameters to be changed to alter appearance and function. This section specifies which parameters need to be set to particular values in order to maintain consistent behaviour across fields and forms.

For all fields:

| Parameter | Value |
| --------- | ------------------------------- |
| BackColor | &H80000005& - Window Background |
| BorderColor | &H80000006& - Window Frame |
| ForeColor | &H80000008& - Window Text |
| BorderStyle | 1 - fmBorderStyleSingle |

For large text input fields:

| Parameter | Value |
| ----------------- | ---- |
| EnterKeyBehaviour | True |
| ScrollBars | True |
| MultiLine | True |

For time fields:

| Parameter | Value |
| --------- | --- |
| MaxLength | 4 |


##### Naming Convention

For the application to function correctly, fields must be named following a specific format. Every field should be named `InputN`, where N is the position of that field on the form. So the first field should be `Input1`, the second `Input2`, and so on. Every field should have a descriptive label, named `LabelN` following the same pattern, such that `Label3` is the label for `Input3`.

##### Updating Form Code
Although most of the application functionality is implemented elsewhere, some changes may need to be made to the code within each actual form object. This code is accessible by double-clicking on a particular form template from the drag and drop view. The form code is separated in `Sub` subroutines, and in the following circumstances, these must be added and edited directly. Subroutines do not need to be arranged in any particular order, but it is highly recommended to organise them in a grouped and logical order.

###### Populating Combo Boxes
If the field being added is a combo box, then code needs to be added to create the options that consist of the fields valid input values. Every form has a `userForm_Initialize()` function, and for a combo box named `Input4`, the following two lines need to be added to this subroutine in order to populate it when a user clicks on it.

```vba
Me.Input4.AddItem "Yes"
Me.Input4.AddItem "No"
```

These lines must come before the final line of the function:

```vba
initialiseForm [form constant name], Me
```

###### Linking Options with Children
If the new field is a combo box that has children, another subroutine must be implemented so that changes to the combo box de/activates its children in real time. If the same `Input4` from the above example is an option field with the form constant `FO_OPT_CTCHECKED`, then the following subroutine would be used:

```vba
Private Sub Input4_Change()
    If UCase(Me.Input1.Value) = "YES" Then
        SetActiveChildFields FO_OPT_CTCHECKED, True
    ElseIf UCase(Me.Input1.Value) = "NO" Then
        SetActiveChildFields FO_OPT_CTCHECKED, False
    End If
End Sub
```

###### Textbox Right Click
For large text boxes, in order enable copy and paste by right clicking the mouse, the following subrouting needs to be added to the form code. For text input `Input2':

```vba
Private Sub Input2_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = vbKeyRButton Then Call ShowTextMenu(Me.Controls("Input2"))
End Sub
```

###### Number-Only Text Input
In some cases, it is desirable to only allow numeric input into a text box. WHere this is the case, the following subroutine must be added to the form code. For text input `Input2`:

```vba
Private Sub Input2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    isInt KeyAscii
End Sub
```

#### Updating Existing Elements

##### Cascading Name Values
Once all new fields (and their accompanying labels) have been added, **all** fields and labels on the form that come after a newly added field must have their names adjusted such that their number matches their new position. So if a form with 10 fields has a new field added at the third position, previous inputs and labels 3 - 10 must then be incremented by one.

##### Updating TabIndex
Along with updating the names of fields, the `TabIndex` must also be adjusted. This is the parameter that determines in which order fields are traversed when the user presses the tab key to jump from field to field. This must be updated when fields are added to ensure users can continue to navigate the form efficiently.

#### Updating Constant Parameters
The second stage of adding fields to forms involves updating the [`ConstParams`](#constparams) module. This involves 5 steps, although some of these might be unnecessary depending on the type of input:

1. Add the field to the list of [form constants](#form-constants)

   If the field is an option field, add it to both 'INPUT' and 'OPT' constant sections. Otherwise, just the 'INPUT' section.
   Use the appropriate prefixes: *`[form abbreviation]_[constant section]_[field name]`*.
   Enumerate the constant appropriately. The number assigned to the constant should match the number in that field's name.
   Update the following field constants for that form accordingly, so that no numbers are duplicated and all constants match their field's number.
   
2. Update the form's [General Parameters](#general-params)

   `NUM_OPTS` and `NUM_FIELDS` should be incremented as appropriate
   `NUM_COMMAS` should be changed if additional line breaks will be needed.
   `FULL_RANGE` and `CELL RANGE` should be updated in order to fit the new size of the form's section in the [`fields`](#updating-fields-excel-sheet) sheet.
   `X_OFFSET` and `Y_OFFSET` should not be changed.
   
3. Add new field to relevent [Parameter Arrays](#parameter-array-declarationinitialisation)
   
   For an example field, `FO_INP_NEWFIELD`:
   
   | Condition | Array Update |
   | --------- | ------------ |
   | `FO_INP_NEWFIELD` is a required field | `requiredInputs(FO_INP_NEWFIELD) = 1` |
   | `FO_INP_NEWFIELD` is a time field | `requiredInputs(FO_INP_NEWFIELD) = 1` |
   | `FO_INP_NEWFIELD` needs spellchecking | `requiredInputs(FO_INP_NEWFIELD) = 1` |
   | `FO_INP_NEWFIELD` end the line of output | `requiredInputs(FO_INP_NEWFIELD) = 1` |
   | `FO_INP_NEWFIELD` is an option field | `requiredInputs(FO_INP_NEWFIELD) = 1` |
   | `FO_INP_NEWFIELD` is child of `FO_INP_PARENT` | `requiredInputs(FO_INP_NEWFIELD) = 1` |
   
#### Updating Excel Sheet
The final stage of adding elements to a particular form is to change the `fields` excel sheet. This sheet is hidden by default, and should be hidden again once changes have been made to it. This sheet defines the titles that get paired to the text from the form in the output that gets sent to the user's clipboard. When a new field is added to a form, that form's title list should be found, and the new field's title should be added in the appropriate location.

### Removing Fields
Removing fields involves most of the same processes as field addition, but reversed:

- Ensure remaining fields are re-positioned to maintain visual cohesion.

- Decrement remaining field and label names to ensure no missing gaps exist.

- Decrement form constants in ConstParams module.

- Update form's General Parameters to reflect new structure.

- Remove references to removed field(s) in Parameter Arrays.

- Remove field title(s) from `field` Sheet, and update range accordingly.

## Deploying for Production
### Rename File Copy
The first stage of production deployment (assuming the working file has had all necessary functional updates made) is to create a copy of the file, and rename it. The naming convention in use as of the writing of this documentation is as follows:

**Working File**: freeform_templates_tool_working.xlsm

**Production File**: freeform_templates_V[*version number*].xlsm

It is also advised to keep a backup working file that is updated whenever a new version is released, so that the working file can always be reverted to the most recent production state. 

### Setup Exit on Close
The second stage of preparing the file for production is to change the code in the `ufrmMain` form to make sure that when the pop up form is closed, the entire application terminates as well. 

The working file keeps the application running when the pop up is closed, allowing the developer to subsequently access the code and make alterations. This is inappropriate for a tool in production, and the following changes in the `UserForm_Terminate()` function prevent this from occurring:
```vba
Private Sub UserForm_Terminate()
   ' Delete this line
    Windows(ThisWorkbook.Name).Visible = True

   ' Add this line
    Application.Quit
End Sub
```
> **Note:** Once this change has been made and the file has been saved and closed, you will no longer be able to access the code to make further edits. As such, it is **extremely important** to make sure to rename a copy of the file before doing so, and ensuring that no further changes need to be made before the file is ready to be put to use.
 
