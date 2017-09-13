![LOGO](https://davecra.files.wordpress.com/2017/07/officejs-dialogs.png?w=698)
# Introduction
The OfficeJS.dialogs library provides simple to use dialogs in OfficeJS/Office Web Add-in (formally called Apps for Office) solutions. The secondary purpose of the library is to help bring some familiarity (from VBA/VB/C#) into OfficeJS development. Currently, the following dialogs types are present:
* [MessageBox](#MessageBox)
* [Alert](#Alert)
* [InputBox](#InputBox)
* [Progress](#Progress)
* [Wait](#Wait)
* [Form](#Form)
* [PrintPreview](#PrintPreview)

# Update History
Current version:  1.0.8
Publish Date:     9/13/2017

This is a breif history of updates that have been applied:

* Version 1.0.1 - support for MessageBox, InputBox and custom Form
* Version 1.0.2 - bug fixes, streamlining code
* Version 1.0.3 - better error handling, cancel code
* Version 1.0.4 - support for Wait form, Progress Dialog, and async updates on MessageBox and ProgressBar while forms are still loaded.
* Version 1.0.5 - removed external image references, converted to inline base64 strings. Bug fixes.
* Version 1.0.6 - converted to classes, bug fixes, updated inline jsdoc documentation, standardized, fixed issues with close dialog - one after another - known bug in OfficeJS/Outlook/OWA
* Version 1.0.7 - bug fixes, code cleanup, documentation added README.md
* Version 1.0.8 - support for PrintPreview, code cleanup , bug fixes

In the following sections each of these will be details with proper usage.

### Installation
To install OfficeJS.dialogs, you can either pull in this repository from GitHub, by cloning it and then importing it into your project, or using the following command in your preferred coding environment with Node installed:

```
npm install officejs.dialogs
```

There is also a CDN here: https://cdn.rawgit.com/davecra/OfficeJS.dialogs/master/dialogs.js

Please note, the CDN has CORS issues with any of the Update() commands below. As such, you will be able to display a Progress dialog, but you will be completely unable to update it (increment). You will also be unable to use the MessageBox.Update() command as well. If you have no need for these commands, by all means, please use the CDN, but be aware of these limitations.

For now, the guidance is to use Node Package Manager (NPM) to import the library into your solution.

### Follow
Please follow my blog for the latest developments on OfficeJS.dialogs. You can find my blog here:

![LOGO](https://davecra.files.wordpress.com/2017/07/blog-icon-large.png?w=20) http://theofficecontext.com

You can use this link to narrow the results only to those posts which relate to this library:

* https://theofficecontext.com/?s=officejs.dialogs
  
![TWITTER](https://davecra.files.wordpress.com/2010/10/tlogo.png?w=20) You can also follow me on Twitter: [@davecra](http://twitter.com/davecra)

![LINKEDIN](https://davecra.files.wordpress.com/2014/02/inbug-60px-r.png?w=20) And also on LinkedIn: [davidcr](https://www.linkedin.com/in/davidcr/)

# MessageBox<a name="MessageBox"></a>
The MessageBox class has the following public methods:
* [Reset()](#MessageBoxReset)
* [Show](#MessageBoxShow)([text],[caption],[buttons],[icon],[withcheckbox],[checkboxtext],[asyncResult],[processupdates])
* [Update](#MessageBoxUpdate)([text],[caption],[buttons],[icon],[withcheckbox],[checkboxtext],[asyncResult])
* [UpdateMessage](#MessageBoxUpdateMessage)([text],[asyncResult])
* [Displayed()](#MessageBoxDisplayed)
* [CloseDialogAsync](#MessageBoxCloseDialog)([asyncresult])

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
          MessageBox.CloseDialogAsync(function() {
            Alert.Show("You said you " + iceCream + " like ice cream, you " +
                        jellyBeans + " like jelly beans, and you " + 
                        kitkat + " like kit kat bars.");
          });
        });
      });
    }, true);
```
This is an example of one MessageBox from the above code:

![MessageBox Dialog](https://davecra.files.wordpress.com/2017/07/messagebox-sample.png?w=600)

### MessageBox.Update()<a name="MessageBoxUpdate"></a>
If you issue a [MessageBox.Show()](#MessageBoxShow) and you set the [processupdated] flag to true, then you can use this method. Otherwise this will fail. What this method does is update a currently displayed messagebox with new information. This has all the same paramaters as the [MessageBox.Show()](#MessageBoxShow) with the exception of the processupdated flag (since the dialog is already setup to allow you to issue updates). You must issue a new callback as well to handle the new updated response. For information on what each paramater does, and defaults, see the [MessageBox.Show()](#MessageBoxShow) method.

### MessageBox.UpdateMessage()<a name="MessageBoxUpdateMessage"></a>
If you issue a [MessageBox.Show()](#MessageBoxShow) and you set the [processupdated] flag to true, then you can use this method. Otherwise this will fail. What this method does is updates just the text message of a currently displayed messagebox. This accepts only the [text] paramater and a [asyncResult] callback. The text and callback are both required. You must issue a new callback as well to handle the new updated response. For information on what each paramater does, and defaults, see the [MessageBox.Show()](#MessageBoxShow) method.

### MessageBox.Displayed()<a name="MessageBoxDisplayed"></a>
This method returns true if a MessageBox dialog is currently being displayed to the user. This is provided in case you wish to verify the dialog is still opened before issuing a [MessageBox.CloseDialog()](#MessageBoxCloseDialog) or [MessageBox.Update()](#MessageBoxUpdate) or [MessageBox.UpdateMessage()](#MessageBoxUpdateMessage).

### MessageBox.CloseDialogAsync()<a name="MessageBoxCloseDialog"></a>
If you issue a [MessageBox.Show()](#MessageBox.Show) and you set the [processupdated] flag to true, then you can use this method to close the dialog. Otherwise this will fail. This will close the currently displayed MessageBox.

**NOTE**: Because of the way Office dialogs work, all dialogs have to be closed asyncronously in order to avoid situations where trying to open a second dialog will fail, because another one is still in the process of being destroyed.

The CloseDialogAsync has the following paramter:
* [**asyncResult**: *function()*] (required) - This is callback is invoked when the dialog is completely closed.

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

![Alert Dialog](https://davecra.files.wordpress.com/2017/07/alert.png?w=500)

### Alert.Displayed()<a name="AlertDisplayed"></a>
This method returns true if an Alert dialog is currently being displayed to the user.

# InputBox<a name="InputBox"></a>
The InputBox class has the follwoing public methods:
* [Reset()](#InputBoxReset)
* [Show](#InputBoxShow)(text,caption,defaultvalue,syncresult)
* [Displayed()](#InputBoxDisplayed)

### InputBox.Reset()<a name="InputBoxReset"></a>
You can issue command each time you are about to request a InputBox dialog to assure everything is reset (as it is in the global space). This resets the InputBox global object so that no previous dialog settings interfere with your new dialog request. You should only use this if you encounter issues.

### InputBox.Show()<a name="InputBoxShow"></a>
This method displays an InputBox to the user with the text message you provide, a caption and a default value. When the user clicks Ok, the result in the callback *asyncresult* will be the text they typed. If the user pressed cancel or clicked the (X) in the upper right of the dialog, the result will be blank. Here are the parameters:
* [**text**: *string*] (required) - this is the question or message you want the user to see in the InputBox.
* [**caption**: *string/256*] (optional) - this is the caption that will appear int eh dialog. The default value will be blank.
* [**defaultvalue**: string] (optional) - this is the default value you want to prepopulate in the textbox od the dialog. The default value is blank.
* [**asyncresult**: *function*(*string*)] (required) - this is the callback with the result from the dialog. If the user pressed cancel, the result is blank.

The following sample asks the user for a subject and then applies the result to the email message:
```javascript
    InputBox.Show("What is the email subject?", "Email Subject", "Default Email Subject", 
      function(result) {
        if(result.length > 0) {
          // in the server service callback
          Office.cast.item.toMessageCompose(Office.context.mailbox.item).subject.setAsync(result,
            function() {
              Alert.Show("The subject has been set to " + result);
            });
        }
    });
```

Here is an example of an InputBox based on the sample code provided above:

![InputBox Dialog](https://davecra.files.wordpress.com/2017/07/inputbox.png?w=500)

### InputBox.Displayed()<a name="InputBoxDisplayed"></a>
This method returns true if an InputBox dialog is currently being displayed to the user.

# Progress<a name="Progress"></a>
The Progress class has the following public methods:
* [Reset()](#ProgressReset)
* [Show](#ProgressShow)([text],[start],[max],[asyncresult],[cancelresult])
* [Update](#ProgressUpdate)([increment],[text])
* [Completed()](#ProgressCompleted)
* [Displayed()](#ProgressDisplayed)

### Progress.Reset()<a name="ProgressReset"></a>
You can issue command each time you are about to request a progress dialog to assure everything is reset (as it is in the global space). This resets the Progress global object so that no previous dialog settings interfere with your new dialog request. You should only use this if you encounter issues.

### Progress.Show()<a name="ProgressShow"></a>
This method will display a progress dialog with the spcified text. You will update the dialog with [Progress.Update()](#ProgressUpdate) to change the value of the progress bar and/or the text in the dialog. Once completed, you will call the [ProgressBar.Complete()](#ProgressComplete) method to close the dialog. You will usually make this call from the *asyncresult* callback. If the user presses cancel at any time while the dialog is loaded, the *cancelresult* callback will be called. Here are the function paramaters:
* [**text**: *string/256*] (optional) - this is the text the user will see. It is limited to 256 characters in length. If none is specified the default is "Please wait..."
* [**start**: *number*] (optional) - this is the starting number for the progress bar. The default is zero (0).
* [**max**: *number*] (optional) - this is the max value of the progress bar. The default is 100.
* [**asyncresult**: *function()*] (optional) - this is the callback when the Progress.Compelte() method is called.
* [**cancelresult**: *function()*] (optional) - this is the callback when the user presses cancel on the dialog.

Here is some example code that display a Progress dialog and then uses a seperate function with a timer to update it until it hits 100%:

```javascript
function dotIt() {
  // display a progress bar form and set it from 0 to 100
  Progress.Show("Please wait while this happens...", 0, 100, function() {
      // once the dialog reached 100%, we end up here
      Progress.CompleteAsync();
      Alert.Show("All done folks!");
    }, function() {
      // this is only going to be called if the user cancels
      Alert.Show("You cancelled the process.");
      // clean up stuff here...

  });
  doProgress();
}

function doProgress() {
  // increment by one, the result that comes back is
  // two pieces of information: Cancelled and Value
  var result = Progress.Update(1);
  // if we are not cancelled and the value is not 100%
  // we will keep going, but in your code you will
  // likely just be incrementing and making sure
  // at each stage that the user has not cancelled
  if(!result.Cancelled && result.Value <= 100) {
    setTimeout(function() {
      // this is only for our example to
      // cause the progress bar to move
      doProgress();
    },100);
  } else if(result.Value >= 100) {
    Progress.Compelte(); // done
  }
}
```

This is an example of a Progress dialog from the code above:

![Progress Dialog](https://davecra.files.wordpress.com/2017/07/progress.png?w=600)

### Progress.Update()<a name="ProgressUpdate"></a>
This method will update the progress. By default if you do not pass any paramaters, the progress bar on the dialog will increment by one. However, you also have the option to change the text and/or the progress increment amount. If you specify an increment of zero (0) and specify new text for the dialog, the text will change, but the dialog will not increment. Here are the parameters:
* [**increment**: *number*] (optional) - this is the amount to increment the progress bar by.
* [**text**: *string/256*] (optional) - you can change the text on the displayed progress bar by issuing new text. It is limited to 256 characters in length.

### Progress.Compelted()<a name="ProgressCompleted"></a>
This method will close the progress dialog. You will usually call this from the *asyncresult* callback setup in the [Progress.Show()](#ProgressShow) method. 

### Progress.Displayed()<a name="ProgressDisplayed"></a>
This method returns true if a Progress dialog is currently being displayed to the user.

# Wait<a name="Wait"></a>
This displays a very simple wait dialog box with a spinning GIF. It has only one option and that is to display the cancel button. Here are the available methods:
* [Show](#WaitShow)([text],[showcancel],[cancelresult])
* [Reset()](#WaitReset)
* [CloseDialogAsync()](#WaitCloseDialog)([asyncResult])
* [Displayed()](#WaitDisplayed)

### Wait.Show()<a name="WaitShow"></a>
This displays a simple wait dialog to the user with a spinning GIF. This dialog will remain open until you issue a Wait.DialogClose(). Here are the parameters:
* [**text**: *string*] (optional) - if text is provided, this is the message the user will see above the spinning GIF. Default is "Please wait..."
* [**showcancel**: *boolean*] (optional) - if this is true, then the user will have the option to cancel the dialog. You will need to provide a [cancelresult] callback in thie case. The default is false.
* [**cancelresult**: *function()*] (optional) - if the showcancel option is enabled, this is required to notify your code that the user pressed cancel. There are not paramters provided in the callback. 

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
        Wait.CloseDialogAsync(function() { });
      }
    });
 ```

This is an example of the Wait dialog from the code above: 

![Wait Dialog](https://davecra.files.wordpress.com/2017/07/wait.png?w=500)

### Wait.Reset()<a name="WaitReset"></a>
You can issue command each time you are about to request a wait dialog to assure everything is reset (as it is in the global space). This resets the Wait global object so that no previous dialog settings interfere with your new dialog request. You should only use this if you encounter issues.

### Wait.CloseDialogAsync()<a name="WaitCloseDialog"></a>
This closes the open Wait dialog.

**NOTE**: Because of the way Office dialogs work, all dialogs have to be closed asyncronously in order to avoid situations where trying to open a second dialog will fail, because another one is still in the process of being destroyed.

The CloseDialogAsync has the following paramter:
* [**asyncResult**: *function()*] (required) - This is callback is invoked when the dialog is completely closed.

### Wait.Displayed()<a name="WaitDisplayed"></a>
This method returns true if a Wait dialog is currently being displayed to the user. 

# Form<a name="Form"></a>
The custom Form allows you to hook up your own HTML page to use OfficeJS.dialogs framework behind the scenes. You Show() your custom form, recieve callbacks with the information you provide from your form and can handle when and how to close the form. The Form object has the following methods:

* [Reset](#FormReset)()
* [Url](#FormUrl)([value])
* [Height](#FormHeight)([value])
* [Width](#FormWidth)([value])
* [HandleClose](#FormHandleClose)([value])
* [AsyncResult](#FormAsyncResult)([value])
* [CloseDialogAsync](#FormCloseDialogAsync)([asyncResult])
* [Displayed](#FormDisplayed)()
* [Show](#FormShow)([url],[height],[width],[handleclose],[asyncresult])

### Form.Reset()<a name="FormReset"></a>
You can issue command each time you are about to request a new custom form dialog to assure everything is reset (as it is in the global space). This resets the Form global object so that no previous dialog settings interfere with your new dialog request. You should only use this if you encounter issues.

### Form.Url()<a name="FormUrl"></a>
This method/property will allow you to set the URL or retrieve the url value as a string. Here is the sole argument:

* [**value**: *string*] (get/set) - If no value is specified, it will return the current URL. If a value is specified, the URL for the custom dialog will be set to this location.

### Form.Height()<a name="FormHeight"></a>
This method/property will allow you to set the height or retrieve the height as a number. Here is the sole argument:

* [**value**: *number*] (get/set) - If no value is specified, it will return the current form height. If a value is specified, the height for the custom dialog will be set to this value.

### Form.Width()<a name="FormWidth"></a>
This method/property will allow you to set the width or retrieve the width value as a number. Here is the sole argument:

* [**value**: *number*] (get/set) - If no value is specified, it will return the current form width. If a value is specified, the width for the custom dialog will be set to this value.

### Form.HandleClose()<a name="FormHandleClose"></a>
This method/property will allow you to set whether OfficeJS.dialogs framework will close the dialog when a messageParent call is recieved, or whther your code will handle the close. Here is the sole argument:

* [**value**: *boolean*] (get/set) - If no value is specified, it will return the current setting. If a value is specified, and that value is **true**, then the OfficeJS.dialogs framework will handle the close of the form for you when any message is recieved via the messageParent call. If the value is set to **false**, you will need to handle the closing of the dialog using the CloseDialogAsync() command.

### Form.AsyncResult()<a name="FormAsyncResult"></a>
This method/property will allow you to set the callback fucntion for the custom form. Here is the sole argument:

* [**value**: *function(string)*] (set only) - Allows you to set the callback function for the close of the form. The callback will recieve a **string** paramater that will return a JSON object formatted as such:
        {
             Error: { },         // Error object
             Result: { },        // JSON from form
             Cancelled: false,   // boolean if form cancelled with X
             Dialog: { }         // A reference to the dialog
        }

### Form.Displayed()<a name="FormDisplayed"></a>
This method returns true if a Form dialog is currently being displayed to the user. 

### Form.Show()<a name="FormShow"></a>
This method allows you to open your own custom form using the framework provided by OfficeJS.dialogs. The form you use must conform in the following ways:

* It must have the following references: 
        <script src="https://ajax.aspnetcdn.com/ajax/jQuery/jquery-1.9.1.min.js" type="text/javascript"></script>
        <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
* It must initialize Office on load:
        Office.initialize = function(reason) { /*... your code here*/ }
* It must issue a callback when you want to update your calling code:
        Office.context.ui.messageParent(JSON.stringify("{'myData':'myValue'}"));

If your dialog does not meet the above requirements, it will not function properly in the OfficeJS.dialogs framework. Here are the parameters for the Show() method:

* [**url**: *string*] (optional) - This is optional only if the Url() value has been set before the Show method is called. Otherwise you will recieve an error. This is the fully qualified URL to your dialog. Please NOTE that cross-domain issues may prevent your code from executing properly if you pass in a URL that is NOT in your current domain.
* [**height**: *number*] (optional) - This will set the height of the form. If no value is specified in either the Height() property or here, the default height will be set to 1 (minimum value).
* [**width**: *number*] (optional) - This will set the width of the form. If no value is specified in either the Width() proeprty or here, the default width will be set to 1 (minium value).
* [**handleclose**: *boolean*] (optional) - This will determine whether the framework will close the form when a messageParent is recieved. If **true** the dialog will be closed automatically when any message is recieved. If **false** you will hae to issue a CloseDialogAsync() when you are ready to close the form.
* [**asyncresult**: *function(string)*] (optional) - This is the callback with the result from the form when your form issues a messageParent() call. The callback recieves a string and your message will be found in the JSON result as defined in the [AsyncResult()](#FormAsyncResult) section above.

Here is a smple of how to use the Form dialog:

```javascript
  Form.Show("/test.html", 20,30, false, function(result) {
    Form.CloseDialogAsync(function() {
      console.log("here");
      Alert.Show("The value is: " + result);
    });
  });
```

Here is a sample of the **test.html** as defined above:

```html
<html>
    <head>
        <script src="https://ajax.aspnetcdn.com/ajax/jQuery/jquery-1.9.1.min.js" type="text/javascript"></script>
        <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
        <script>
            Office.initialize = function(reason) { 
                $(document).ready(function () {
                    $("#okButton").click(function() {
                        // notify the parent - FunctionFile
                        Office.context.ui.messageParent(JSON.stringify(
                            { 
                                "FibbyGibber": "rkejfnlwrjknflkerjnf",
                                "DoDaDay": "Hahahaha",
                                "Message" : "My custom message." 
                            }));
                    });
                });
            };
        </script>
    </head>
    <body>
        Click the button <br/>
        <button id="okButton">Ok</button>
    </body>
</html>
```

Here is what the above dialog look like when issued:

![Form Dialog](https://davecra.files.wordpress.com/2017/07/formdialog.png?w=300)

Here are the JSON results:

```json
{
  "Error":{},
  "Result":"{
              'FibbyGibber':'rkejfnlwrjknflkerjnf',
              'DoDaDay':'Hahahaha',
              'Message':'My custom message.'
             }",
  "Cancelled":false
}
```

# PrintPreview<a name="PrintPreview"></a>
The PrintPreview form allows you to send any HTML to the dialog to de displayed in the form (via iframe). In the dialog the user will have the option to cancel, or to Print. When the user clicks Print a new window will be opened, the contents of the frame will be placed in the window and it will be printed. The PrintPreview object has the following methods:

* [Reset](#PrintReset)()
* [Displayed](#PrintDisplayed)()
* [Show](#PrintShow)([html],[cancelresult])

### PrintPreview.Reset()<a name="PrintReset"></a>
You can issue command each time you are about to request a PrintPreview dialog to assure everything is reset (as it is in the global space). This resets the PrintPreview global object so that no previous dialog settings interfere with your new dialog request. You should only use this if you encounter issues.

### PrintPreview.Displayed()<a name="PrintDisplayed"></a>
This method returns true if a PrintPreview dialog is currently being displayed to the user. 

### PrintPreview.Show()<a name="PrintShow"></a>
This method opens the PrintPreview dialog using the HTML by OfficeJS.dialogs. Here are the parameters for the Show() method:

* **html**: *string* (required) - This is the html you want to display in the dialog. You cna either get this from the document/body/selection of the Office item you are using or submit your own custom HTML if printing a custom form, for example.
* [**cancelresult**: *function()*] (optional) - This is the callback if the user presses cancel. It is not required.

Here is a smple of how to use the PrintPreview dialog:

```javascript
  // this example takes the currently composed email message in Outlook,
  // grabs its body HTML and then displays it in the Print Preview dialog.
  var mailItem = Office.cast.item.toItemCompose(Office.context.mailbox.item);
  mailItem.saveAsync(function(asyncResult) {
    var id = asyncResult.id;
    mailItem.body.getAsync(Office.CoercionType.Html, { asyncContext: { var3: 1, var4: 2 } }, function(result) {
      var html = result.value;
      PrintPreview.Show(html, function() {
        Alert.Show("Print cancelled");
      });
    });
  });
```

![PrintPreview Dialog](https://davecra.files.wordpress.com/2017/09/print.png?w=500)