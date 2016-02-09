# Outlook 2016 Add-in Bug Demo

In Outlook 2016 15.18 on Mac, it appears that local storage is removed every time the plugin is opened.  This makes it very difficult to persist user data like tokens between sessions.

### Setup

To run the demo, first `npm install && bower install` to install dependencies.
Run `gulp serve-static` to start the server at localhost:8443.
Install the add-in manifest in the project to Outlook.

### Behavior

On each visit, the add-in will check local storage for the key foo and if there is a value, it will display that value and then remove it from LS.  If there is not a value, the value bar will be added for foo.  When you open and close the plugin, it should toggle between `no foo` and `bar`.

On Outlook 2016 15.18 on Mac, it removes local storage values every time, so you'll only see `no foo`.
