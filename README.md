![LOGO](https://davecra.files.wordpress.com/2017/07/officejs-dialogs.png?w=698)
# Introduction
The OfficeJS.dialogs library provides simple to use dialogs in OfficeJS/Office Web Add-in (formally called Apps for Office) solutions. The secondary purpose of the library is to help bring some familiarity (from VBA/VB/C#) into OfficeJS development. Currently, the following dialogs types are present:
* [MessageBox](#MessageBox)
* [Alert](#Alert)
* [InputBox](#InputBox)
* [Progress](#Progress)
* [Wait](#Wait)

In the following sections each of these will be details with proper usage.

# MessageBox<a name="MessageBox"></a>
The MessageBox class has four public methods:
* [Reset()](#MessageBoxReset)
* [Show](#MessageBoxShow)([text],[caption],[buttons],[icon],[withcheckbox],[checkboxtext],[asyncResult],[processupdates])
* [Update](#MessageBoxUpdate)([text],[caption],[buttons],[icon],[withcheckbox],[checkboxtext],[asyncResult])
* [UpdateMessage](#MessageBoxUpdateMessage)([text],[asyncResult])
* [Displayed()](#MessageBoxDisplayed)
* [CloseDialog()](#MessageBoxCloseDialog)

### MessageBox.Reset()<a name="MessageBoxReset"></a>
You can issue command each time you are about to request a messagebox dialog to assure everything is reset (as it is in the global space). This resets the MessageBox global object so that no previous dialog settings interfere with your new dialog request. You should only use this if you encounter issues.

### MessageBox.Show()<a name="MessageBoxShow"></a>
The Show method will display a MessageBox dialog with a caption, a message, a selection of buttons (OK, Cancel, Yes, No, Abort, and Retry) and an icon (Excalation, Asterisk, Error, Hand, Information, Question, Stop, Warning), an optional checkbox with its own text messge, a callback when the user pressses any of the buttons and an option to keep the dialog open until you issue a DialogClose(). The following paramaters are used in this method:
* [**text**: *string*] (required) - this is the main message you want to display in the dialog.
* [**caption**: *string*] (optional) - this is the caption to appear above the main message. Default is blank.
* [**buttons**: *MessageboxButtons*] (optional) - this is a member of the MessageBoxButtons enumeration. You can pick between: OkOnly, OkCancel, YesNo, YesNoCancel, RetryCancel, AbortRetryCancel. Default is OkOnly.
* [**icon**: *MessageBoxIcons*] (optional) - this is a member of the MessageBoxIcons enumeration. Default is None.
* [**withcheckbox**: *boolean*] (optional) - if this is enabled a checkbox will appear at the bottom of the form. Should be used in conjunction with the [checkboxtext] parameter below. This is useful for providing an option like: "Do not show this message again." The default is false.
* [**checkboxtext**: *string*] (optional) - if the [withcheckbox] option is enabled, this is the text that will appear to the right of the checkbox. Default is blank.
* [**asynResult**: *function*(*string*,*boolean*)] (required) - this is the callback function that returns which button the user pressed (as a string) and whether then user checked the checkbox (is the [withcheckbox] option is enabled). If the user presses the (X) at the top right of the dialog, "CANCEL" will be returned.
* [**processupdates**: *boolean*] (optional) - if this option is true, the dialog will remain open after the user presses a button. The message will be sent to the callback but the dialog will continue to remain. This is useful if you have a series of questions to ask the user in rapid succession, rather than closing the dialog and reopening it each time (returning control temporarily to the Office application), you can issue UpdateMessage() or Update() to change the message, buttons, caption, icon and callback. If this option is true, you are responsible for closing the dialog when complete by issuing a MEssageBox.CloseDialog() command. The default is false. This means that when the user presses any button the dialog closes.

```javascript
  MessageBox.Show("Do you like icecream?", "Questionaire", MessageBoxButtons.YesNo, 
      MessageBoxIcons.Question, false, null,function(buttonFirst) {
      /** @type {string} */
      var iceCream = (buttonFirst == "Yes" ? "do" : "dont");
      MessageBox.UpdateMessage("Do you like Jelly Beans?", function(buttonSecond) {
        /** @type {string} */
        var jellyBeans = (buttonSecond == "Yes" ? "do" : "dont");
        MessageBox.UpdateMessage("Do you like Kit Kat bars?", function(buttonThird) {
          /** type {string} */
          var kitkat = (buttonThird == "Yes" ? "do" : "dont");
          MessageBox.CloseDialog();
          Alert.Show("You said you " + iceCream + " like ice cream, you " +
                      jellyBeans + " like jelly beans, and you " + 
                      kitkat + " like kit kat bars.");
        });
      });
    }, true);
```
This is an example of one MessageBox from the above code:

<img src=https://davecra.files.wordpress.com/2017/07/messagebox-sample.png width=600>

### MessageBox.Update()<a name="MessageBoxUpdate"></a>
If you issue a [MessageBox.Show()](#MessageBoxShow) and you set the [processupdated] flag to true, then you can use this method. Otherwise this will fail. What this method does is update a currently displayed messagebox with new information. This has all the same paramaters as the [MessageBox.Show()](#MessageBoxShow) with the exception of the processupdated flag (since the dialog is already setup to allow you to issue updates). You must issue a new callback as well to handle the new updated response. For information on what each paramater does, and defaults, see the [MessageBox.Show()](#MessageBoxShow) method.

### MessageBox.UpdateMessage()<a name="MessageBoxUpdateMessage"></a>
If you issue a [MessageBox.Show()](#MessageBoxShow) and you set the [processupdated] flag to true, then you can use this method. Otherwise this will fail. What this method does is updates just the text message of a currently displayed messagebox. This accepts only the [text] paramater and a [asyncResult] callback. The text and callback are both required. You must issue a new callback as well to handle the new updated response. For information on what each paramater does, and defaults, see the [MessageBox.Show()](#MessageBoxShow) method.

### MessageBox.Displayed()<a name="MessageBoxDisplayed"></a>
This method returns true if a MessageBox dialog is currently being displayed to the user. This is provided in case you wish to verify the dialog is still opened before issuing a [MessageBox.CloseDialog()](#MessageBoxCloseDialog) or [MessageBox.Update()](#MessageBoxUpdate) or [MessageBox.UpdateMessage()](#MessageBoxUpdateMessage).

### MessageBox.CloseDialog()<a name="MessageBoxCloseDialog"></a>
If you issue a [MessageBox.Show()](#MessageBox.Show) and you set the [processupdated] flag to true, then you can use this method. Otherwise this will fail. This will close the currently displayed MessageBox.

# Alert<a name="Alert"></a>
The alert dialog is the simplest of all. It has only two methods: Show() and Displayed(). Here are the details:
* [Show](#AlertShow)([text], [asyncResult])
* [Displayed()](#AlertDisplayed)

### Alert.Show()<a name="AlertShow"></a>
The Alert.Show() method will display a simple dialog with only up to 256 characters of text and an OK button. When the user presses OK, the dialog is dismissed. When the user presses OK, the callabck [asyncResult] is called. Here are the details on the paramters.
* [**text**: *string/256*] (required) - This is the message you wish to display to the user. It is trimmed at 256 characters in length.
* [**asynResult**: *function()*] (required) - This is the callback which is invoked when the user presses the OK button or clicks the (X) in the upper right of the dialog. There are no paramters in the callback.

```javascript
    const BAD_SUBJECT_CONTENT = "BAD";
    Office.cast.item.toMessageCompose(Office.context.mailbox.item).subject.getAsync(function(result) {
      /** @type {string} */
      var subject = result.value;
      if(subject.indexOf(BAD_SUBJECT_CONTENT) > 0) {
        Alert.Show("You have invalid content in the email subject.");
      }
    });
```

This is an example of the Alert dialog from the code above:

<img src="https://davecra.files.wordpress.com/2017/07/messagebox-sample.png" width=500>

### Alert.Displayed()<a name="AlertDisplayed"></a>
This method returns true if an Alert dialog is currently being displayed to the user.

# InputBox<a name="InputBox"></a>
This section is TDB.

# Progress<a name="Progress"></a>
This section is TDB.

# Wait<a name="Wait"></a>
This displays a very simple wait dialog box with a spinning GIF. It has only one option and that is to display the cancel button. Here are the available methods:
* [Show](#WaitShow)([text],[showcancel],[cancelresult])
* [CloseDialog()](#WaitCloseDialog)
* [Displayed()](#WaitDisplayed)

### Wait.Show()<a name="WaitShow"></a>
This displays a simple wait dialog to the user with a spinning GIF. This dialog will remain open until you issue a Wait.DialogClose(). Here are the parameters:
* [text:string] (optional) - if text is provided, this is the message the user will see above the spinning GIF. Default is "Please wait..."
* [showcancel:boolean] (optional) - if this is true, then the user will have the option to cancel the dialog. You will need to provide a [cancelresult] callback in thie case. The default is false.
* [cancelresult:function()] (optional) - if the showcancel option is enabled, this is required to notify your code that the user pressed cancel. There are not paramters provided in the callback. 

Here is an example of how to use the Wait dialog:
```javascript
    var cancelled = false;
    Wait.Show(null, true, function() {
      Alert.Show("You have cancelled the process.");
      cancelled = true;
    });
    // change the subject after getting it from the server service
    getSubjectFromServerService(function(result) {
      if(!cancelled) {
        // in the server service callback
        Office.cast.item.toMessageCompose(Office.context.mailbox.item).subject = result.value;
        Wait.CloseDialog();
      }
    });
 ```

This is an example of the Wait dialog from the code above: 

<img src="https://davecra.files.wordpress.com/2017/07/messagebox-sample.png" width=500>

### Wait.CloseDialog()<a name="WaitCloseDialog"></a>
This closes the open Wait dialog.

### Wait.Displayed()<a name="WaitDisplayed"></a>
This method returns true if a Wait dialog is currently being displayed to the user. 
