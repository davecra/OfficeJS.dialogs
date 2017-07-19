![LOGO](https://davecra.files.wordpress.com/2017/07/officejs-dialogs.png?w=698)
# Introduction
The OfficeJS.dialogs library provides simple to use dialogs in OfficeJS/Office Web Add-in (formally called Apps for Office) solutions. The secondary purpose of the library is to help bring some familiarity (from VBA/VB/C#) into OfficeJS development. Currently, the following dialogs types are present:
* [MessageBox](#MessageBox)
* [Alert](#Alert)
* [InputBox](#ImputBox)
* [Progress](#Progress)
* [Wait](#Wait)

In the following sections each of these will be details with proper usage.

# MessageBox<a name="MessageBox"></a>
The MessageBox class has four public methods:
* [Reset()](#MessageBox.Reset())
* [Show](#MessageBox.Show())([text],[caption],[buttons],[icon],[withcheckbox],[checkboxtext],[asyncResult],[processupdates])
* [Update](#MessageBox.Update())([text],[caption],[buttons],[icon],[withcheckbox],[checkboxtext],[asyncResult])
* [UpdateMessage](#MessageBox.UpdateMessage()]([text],[asyncResult])
* [Displayed()](#MessageBox.Displayed())
* [CloseDialog()](#MessageBox.CloseDialog()]

### MessageBox.Reset()
You should issue this command each and every time you are about to request a messagebox dialog. This reset the environment and verifies that no previous dialogs or settings interfere with your new dialog request.

### MessageBox.Show()
The Show method will display a MessageBox dialog with a caption, a message, a selection of buttons (OK, Cancel, Yes, No, Abort, and Retry) and an icon (Excalation, Asterisk, Error, Hand, Information, Question, Stop, Warning), an optional checkbox with its own text messge, a callback when the user pressses any of the buttons and an option to keep the dialog open until you issue a DialogClose(). The following paramaters are used in this method:
* [text:string] (required) - this is the main message you want to display in the dialog.
* [caption:string] (optional) - this is the caption to appear above the main message. Default is blank.
* [buttons:MessageboxButtons] (optional) - this is a member of the MessageBoxButtons enumeration. You can pick between: OkOnly, OkCancel, YesNo, YesNoCancel, RetryCancel, AbortRetryCancel. Default is OkOnly.
* [icon:MessageBoxIcons] (optional) - this is a member of the MessageBoxIcons enumeration. Default is None.
* [withcheckbox:boolean] (optional) - if this is enabled a checkbox will appear at the bottom of the form. Should be used in conjunction with the [checkboxtext] parameter below. This is useful for providing an option like: "Do not show this message again." The default is false.
* [checkboxtext:string] (optional) - if the [withcheckbox] option is enabled, this is the text that will appear to the right of the checkbox. Default is blank.
* [asynResult:function(string,boolean)] (required) - this is the callback function that returns which button the user pressed (as a string) and whether then user checked the checkbox (is the [withcheckbox] option is enabled). If the user presses the (X) at the top right of the dialog, "CANCEL" will be returned.
* [processupdates:boolean] (optional) - if this option is true, the dialog will remain open after the user presses a button. The message will be sent to the callback but the dialog will continue to remain. This is useful if you have a series of questions to ask the user in rapid succession, rather than closing the dialog and reopening it each time (returning control temporarily to the Office application), you can issue UpdateMessage() or Update() to change the message, buttons, caption, icon and callback. If this option is true, you are responsible for closing the dialog when complete by issuing a MEssageBox.CloseDialog() command. The default is false. This means that when the user presses any button the dialog closes.

### MessageBox.Update()
If you issue a MessageBox.Show() and you set the [processupdated] flag to true, then you can use this method. Otherwise this will fail. What this method does is update a currently displayed messagebox with new information. This has all the same paramaters as the MessageBox.Show() with the exception of the processupdated flag (since the dialog is already setup to allow you to issue updates). You must issue a new callback as well to handle the new updated response. For information on what each paramater does, and defaults, see the MessageBox.Show() method.

### MessageBox.UpdateMessage()
If you issue a MessageBox.Show() and you set the [processupdated] flag to true, then you can use this method. Otherwise this will fail. What this method does is updates just the text message of a currently displayed messagebox. This accepts only the [text] paramater and a [asyncResult] callback. The text and callback are both required. You must issue a new callback as well to handle the new updated response. For information on what each paramater does, and defaults, see the MessageBox.Show() method.

### MessageBox.Displayed()
This method returns true of the MessageBox is currently being displayed to the user. This is provided in case you wish to verify the dialog is still opened before issuing a CloseDialog() or Update()/UpdateMessage().

### MessageBox.CloseDialog()
If you issue a MessageBox.Show() and you set the [processupdated] flag to true, then you can use this method. Otherwise this will fail. This will close the currently displayed MessageBox.

# Alert<a name="Alert"></a>
The alert dialog is the simplest of all. It has only two methods: Show() and Displayed(). Here are the details:
* Show([text], [asyncResult])
* Displayed()

### Alert.Show()
The Alert.Show() method will display a simple dialog with only up to 256 characters of text and an OK button. When the user presses OK, the dialog is dismissed. When the user presses OK, the callabck [asyncResult] is called. Here are the details on the paramters.
* [text:string/256] (required) - This is the message you wish to display to the user. It is trimmed at 256 characters in length.
* [asynResult:function()] (required) - This is the callback which is invoked when the user presses the OK button or clicks the (X) in the upper right of the dialog. There are no paramters in the callback.

### Alert.Displayed()
This method returns true of an Alert is currently being displayed to the user.

