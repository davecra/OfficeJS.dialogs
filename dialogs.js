/*!
 * dialogs JavaScript Library v1.0.3
 * http://theofficecontext.com
 *
 * Copyright David E. Craig and other contributors
 * Released under the MIT license
 * https://tldrlegal.com/license/mit-license
 *
 * Date: 2017-06-19T15:44EST
/**
 * The global messagebox object for single use of displaying
 * a Message Box in the Office client. Use the Show() method. 
 * @type {msgbox} 
 * */
var MessageBox = new msgbox();
/**
 * The global inputbox object for single use of displaying
 * a Input Box in the Office client. Use the Show() method. 
 * @type {ibox} 
 * */
var InputBox = new ibox();
/**
 * The global form object for single use of displaying
 * a custom form in the Office client. Use the Show() method. 
 * @type {form} 
 * */
var Form = new form();
/**
 * An enum of Message Box Button types
 * @readonly
 * @typedef {string} MessageBoxIcons
 * @enum {MessageBoxIcons} 
 */ 
var MessageBoxIcons = {
    Asterisk: "Asterisk",           // Warning
    Error: "Error",                 // Stop
    Exclamation: "Exclamation",     // Warning
    Hand: "Hand",                   // Stop
    Information: "Information",     // Information
    None: "None",                   // none
    Question: "Question",           // Question
    Stop: "Stop",                   // Stop
    Warning: "Warning"              // Warning
};
/**
 * An enum of Message Box Button types
 * @readonly
 * @typedef {string} MessageBoxButtons
 * @enum {MessageBoxButtons} 
 */ 
var MessageBoxButtons = {
    Ok: "Ok",
    OkCancel: "OkCancel",   
    YesNo: "YesNo",
    YesNoCancel: "YesNoCancel",
    RetryCancel: "RetryCancel",
    AbortRetryCancel: "AbortRetryCancel" 
};
/**
 * A class for creating message boxes in OfficeJS Web Addins
 * @class
 */
function msgbox() {
    /** @type {object} */
    var dialog;
    /** @type {function(string,boolean)} */
    var callback;
    /**
     * Shows the message box, with the provided parameters
     * @param {string} text The message to be shown in the message box
     * @param {string} [caption] The caption on the top of the message box
     * @param {MessageBoxButtons} [buttons] The buttons to be displayed on the message box, of type MessageBoxButtons
     * @param {MessageBoxIcons} [icon] The icon to show on the message box, of type MessageBoxIcons
     * @param {boolean} [withcheckbox] Enables a checkbox on the message box below the buttons
     * @param {string} [checkboxtext] The message to show on the message box checkbox
     * @param {function(string, boolean)} asyncResult Results after the message box is mismissed: 
     *                                                     - String result of the button pressed 
     *                                                     - And boolean is the checkbox was checked
     */
    this.Show = function(text, caption, buttons, icon, withcheckbox, checkboxtext, asyncResult) {
        try {
            // verify
            if(text == null || text.length == 0) {
                throw("No text for messagebox. Cannot proceeed.");
            }
            if(caption == null) caption = "";
            if(buttons == null) buttons = MessageBoxButtons.Ok;
            if(icon == null) icon = MessageBoxIcons.None;
            if(withcheckbox == null) withcheckbox = false;
            if(checkboxtext == null) checkboxtext = "";
            if(asyncResult == null) {
                throw("No callback specified for MessageBox. Cannot proceed.");
            }
            /**
             * Define the settings for the HTML form 
             * @type {object} 
             * */
            var settings = { Text: text, Caption: caption, Buttons: buttons,
                            Icon: icon, WithCheckbox: withcheckbox, 
                            CheckBoxText: checkboxtext, DialogType: "msg" };
            // set the storage item for the dialog form
            localStorage.setItem("dialogSettings", JSON.stringify( settings ));
            // set the callback
            callback = asyncResult;
            var msgWidth = 40;
            var msgHeight = 30; // with checkbox
            if(!withcheckbox) {
                msgHeight = 26; // without
            }
            // show the dialog
            Office.context.ui.displayDialogAsync(getUrl() + "dialogs.html",
                    { height: msgHeight, width: msgWidth, displayInIframe: isOfficeOnline() },
                    function (result) {
                        dialog = result.value;
                        dialog.addEventHandler(Office.EventType.DialogEventReceived, function(arg) {
                            processMsgBoxMessage(arg);
                        });
                        dialog.addEventHandler(Office.EventType.DialogMessageReceived, function(arg) {
                            processMsgBoxMessage(arg);
                        });
                    });
        } catch (e) {
            console.log(e);
        }
    }
    /**
     * Resets the MessageBox object for reuse
     */
    this.Reset = function() {
        try {
            MessageBox = new msgbox();
        } catch(e) {
            console.log(e);
        }
    };
    /**
     * Processes the message from the dialog HTML
     * @param {string | string} arg An object with the results
     */
    function processMsgBoxMessage(arg) {
        try {
            /** @type {string} */
            var button = "";
            /** @type {boolean} */
            var checked = false;
            // process any errors first if there is one and then exit this function, do not
            // process the message. The main one we care about is the user pressing the (X)
            // to close the form. We want to make sure we reset everything.
            /** @type {string} */
            var result = dialogErrorCheck(arg.error);
            if(result == "CANCELLED") {
                // user clicked the (X) to close the dialog
                button = "Cancel";
                checked = false;
            } else if (result == "NOERROR") {
                button = JSON.parse(arg.message).Button;
                checked = JSON.parse(arg.message).Checked;
            } else {
                button = JSON.stringify({ Error: result });
            }
            // close the dialog
            dialog.close();
            // return
            callback(button, checked);
        } catch (e) {
            console.log(e);
        }
    }
    return this;
}
/**
 * Shows the input box, with the provided parameters
 * @class
 */
function ibox(text, caption, defaultValue, asyncResult) {
    /** @type {object} */
    var dialog;
    /** @type {function(string,boolean)} */
    var callback;
    /**
     * Shows the input box, with the provided parameters
     * @param {string} text The message to be shown in the input box
     * @param {string} [caption] The caption on the top of the input box
     * @param {string} [defaultvalue] The default value to be provided
     * @param {function(string)} asyncResult Results after the input box is mismissed. If the
     *                                       returned string is empty, then the user pressed
     *                                       cancel. Otherwise it contains the value the user
     *                                       typed into the form
     */
    this.Show = function(text, caption, defaultvalue, asyncResult) {
        try {
                        // verify
            if(text == null || text.length == 0) {
                throw("No text for InputBox. Cannot proceeed.");
            }
            if(caption == null) caption = "";
            if(defaultvalue == null) defaultvalue = "";
            if(asyncResult == null) {
                throw("No callback specified for InputBox. Cannot proceed.");
            }
            /**
             * Define the settings for the HTML form 
             * @type {object} 
             * */
            var settings = { Text: text, Caption: caption, Buttons: MessageBoxButtons.OkCancel,
                                Icon: MessageBoxIcons.Question, WithCheckbox: false, 
                                CheckBoxText: "", DialogType: "input", DefaultValue: defaultvalue };
            
            // set the storage item for the dialog form
            localStorage.setItem("dialogSettings", JSON.stringify( settings ));
            // set the callback
            callback = asyncResult;
            var msgWidth = 40;
            var msgHeight = 25;
            // show the dialog
            Office.context.ui.displayDialogAsync(getUrl() + "dialogs.html",
                    { height: msgHeight, width: msgWidth, displayInIframe: isOfficeOnline() },
                    function (result) {
                        dialog = result.value;
                        dialog.addEventHandler(Office.EventType.DialogEventReceived, function(arg) {
                            processInputBoxMessage(arg);
                        });
                        dialog.addEventHandler(Office.EventType.DialogMessageReceived, function(arg) {
                            processInputBoxMessage(arg);
                        });
                    });
        } catch (e)
        {
            console.log(e);
        }
    }
    /**
     * Resets the MessageBox object for reuse
     */
    this.Reset = function() {
        try {
            InputBox = new ibox();
        } catch (e) {
            console.log(e);
        }
    };
    /**
     * Processes the message from the dialog HTML
     * @param {string | string} arg An object with the results
     */
    function processInputBoxMessage(arg) {
        try {
            /** @type {string} */
            var text = "";
            // process any errors first if there is one and then exit this function, do not
            // process the message. The main one we care about is the user pressing the (X)
            // to close the form. We want to make sure we reset everything.
            /** @type {string} */
            var result = dialogErrorCheck(arg.error);
            if(result == "CANCELLED") {
                // user clicked the (X) to close the dialog
                text = "";
            } else if(result == "NOERROR") {
                text = JSON.parse(arg.message).Text;
            } else {
                text = JSON.stringify({ Error: result });
            }
            // close the dialog
            dialog.close();
            // return
            callback(text);
        } catch (e) {
            console.log(e);
        }
    }
    return this;
}
/**
 * This class helps create a user form in a dialog
 * @class
 */
function form() {
    /** @type {object} */
    var dialog;
    /**
     * Internal referenced values 
     * @type { } 
     * */
    var value = {
        Url: "",
        Height: 20,         // default
        Width: 30,          // default
        HandleClose: true,  // default
        AsyncResult: { },
        Dialog: { }
    }
    /**
     * Property: Get/Set: The url for the form
     * @param {string} [item] SETTER: The url item you want to set
     * @returns {string} GETTER: If item is null, will return the url
     */
    this.Url = function(item) {
        try { 
            if(item == null) {
                return value.Url;
            } else {
                // the user can specify an folder off the root
                if(item.indexOf("https://") <= 0 && !item.startsWith("/")) {
                    this.Url = getUrl() + item;
                } else if(url.startsWith("/")) {
                    this.Url = getUrl(true) + item;
                } else {
                    this.Url = item; // a fully qualified url
                }
            }
        }
        catch(e) {
            console.log(e);
            return null;
        }
    }
    /**
     * Property: Get/Set: The Height of the form
     * @param {Number} [item] SETTER: The height you want the form to be
     * @returns {Number} GETTER: If item is null, returns the height of the form
     */
    this.Height = function(item) {
        try {
            if(item == null) {
                return value.Height;
            } else {
                value.Height = item;
            }
        } catch(e) {
            console.log(e);
            return null;
        }
    }
    /**
     * Property: Get/Set: The Width of the form
     * @param {Number} [item] SETTER: The width you want the form to be
     * @returns {Number} GETTER: If the item is null, returns the width fo the form
     */
    this.Width = function(item) {
        try {
            if(item == null) {
                return value.Width;
            } else {
                value.Width = item;
            }
        } catch(e) {
            console.log(e);
            return null;
        }
    }
    /**
     * Property: Get/Set: If true the form will close when a message is recieved.
     *                    If false, the caller will have to handle the dialog.close();
     * @param {boolean} [item] SETTER: Sets whether the form will close when a message is recieved
     * @returns {boolean} GETTER: The value of whether the form will close when it recieves a message
     */
    this.HandleClose = function(item) {
        try {
            if(item == null) {
                return value.HandleClose;
            } else {
                value.HandleClose = item;
            }
        } catch (e) {
            console.log(e);
            return null;
        }
    }
    /**
     * Property: Set Only: Sets the callback function only
     * @param {function(string)} The callback function
     */
    this.AsyncResult = function(item) {
        try {
            value.AsyncResult = item;
        } catch(e) {
            console.log(e);
            return null;
        }
    }
    /**
     * Method: Will close the dialog
     */
    this.DialogClose = function() {
        try { 
            dialog.close();
        } catch (e) {
            console.log(e);
            return null;
        }
    }
    /**
     * Shows a form, with the provided parameters
     * @param {string} [url] The url to the form
     * @param {number} [height] The height of the form
     * @param {number} [width] The width of the form
     * @param {boolean} [handleclose] If true, when the form is dismissed the dialog will be closed.
     *                                Otherwise, it is left open and the caller will have to handle
     *                                the dialog.close()
     * @param {function(string)} [asyncResult] Results after the form is dismissed. The 
     *                                         result will be a JSON object like this:
     *                                         {
     *                                              Error: { },         // Error object
     *                                              Result: { },        // JSON from form
     *                                              Cancelled: false,   // boolean if form cancelled with X
     *                                              Dialog: { }         // A reference to the dialog
     *                                         }
     */
    this.Show = function(url, height, width, handleclose, asyncResult) {
        try {
            // set the callback
            if(asyncResult) {
                value.AsyncResult = asyncResult;
            }
            if(height && width) {
                // set the other values
                value.Height = height;
                value.Width = width;
            }
            // set the url
            if(url) {
                // the suer can specify an folder off the root
                if(url.indexOf("https://") <= 0 && !url.startsWith("/")) {
                    value.Url = getUrl() + url;
                } else if(url.startsWith("/")) {
                    // add the host name, assuming we have a full relative path
                    // from the host name and then remove the leading /
                    value.Url = getUrl(true) + url.replace("/","");
                } else {
                    value.Url = url; // a fully qualified url
                }
            } 
            // handle close
            if(handleclose != null) {
                value.HandleClose = handleclose;
            }
            // verify
            if(value.Url == null || value.Url.length == 0) {
                throw("No url specified for form. Cannot proceed.");
            }
            if(!value.AsyncResult) {
                throw("No callback specified for form. Cannot proceed.");
            }
            // show the dialog
            Office.context.ui.displayDialogAsync(value.Url,
                    { height: value.Height, width: value.Width, displayInIframe: isOfficeOnline() },
                    function (result) {
                        dialog = result.value;
                        dialog.addEventHandler(Office.EventType.DialogEventReceived, function(arg) {
                            processFormMessage(arg);
                        });
                        dialog.addEventHandler(Office.EventType.DialogMessageReceived, function(arg) {
                            processFormMessage(arg);
                        });
                    }
                );   
        } catch (e) {
            console.log(e);
        } 
    }
    /**
     * Resets the Form object for reuse
     */
    this.Reset = function() {
        try {
            Form = new form();
        } catch(e) {
            console.log(e);
        }
    };
    /**
     * Processes the message from the dialog HTML
     * @param {string | string} arg An object with the results
     */
    function processFormMessage(arg) {
        try {
            /**@type { } */
            var returnVal = {
                Error: { },             // Error object
                Result: { },            // JSON from form
                Cancelled: false,       // boolean if formed cancelled with X
            };
            // process any errors first if there is one and then exit this function, do not
            // process the message. The main one we care about is the user pressing the (X)
            // to close the form. We want to make sure we reset everything.
            /** @type {string} */
            var result = dialogErrorCheck(arg.error);
            if(result == "CANCELLED") {
                // user clicked the (X) to close the dialog
                returnVal.Cancelled = true;
            } else if(result == "NOERROR") {
                returnVal.Result = arg.message;
            } else {
                // an error occurred
                returnVal.Error = result;
            }
            // close the dialog
            if(value.HandleClose) {
                dialog.close();
            } 
            // return
            value.AsyncResult(JSON.stringify(returnVal));
        } catch (e) {
            console.log(e);
        }
    }
}
/* HELPER FUNCTIONS */
/**
 * Gets the URL of this JS file so we can then grab the dialog html
 * that will be in the same folder
 * @param {boolean} [convert] appends the server name to a relative path
 *                            such as (/folder/pages/page.html) will become
 *                            https://server/folder/pages/page.html 
 * @returns {string} The URL
 */
function getUrl(convert){
    try {
        // /** @type {string} */
        var url = getScriptURL(); // document.location.href;
        if(convert) {
            url = "https://" + document.location.host + "/";
        }
        /** @type {number} */
        var pos = url.lastIndexOf("/");
        url = url.substring(0, pos);
        if(!url.endsWith("/")) {
            url += "/";
        }
        return url;
    } catch(e) {
        console.log(e);
        return null;
    }
}
/**
 * Returns whether the platform is OffOnline
 * @returns {boolean} True if it is OfficeOnline
 */
function isOfficeOnline() {
    /**
     * Check to see if we are in full client or not
     * @type {string}
     */
    var platform = Office.context.platform;
    if(platform == "OfficeOnline") {
        return true;
    } else {
        return false;
    }
}
/**
 * Returns the error details if there is an error number
 * @param {Number} error 
 * @returns {string} Retruns an error message or NOERROR if there is none, 
 *                   or CANCELLED if the dialog was cancelled
 */
function dialogErrorCheck(error) {
    if(error == 12006) {
        return "CANCELLED";
    } else if(error > 0 ) {
        return error.message;
    } else {
        return "NOERROR";
    }
}

var getScriptURL = (function() {
    var scripts = document.getElementsByTagName('script');
    var index = scripts.length - 1;
    var myScript = scripts[index];
    return function() { return myScript.src; };
})();
