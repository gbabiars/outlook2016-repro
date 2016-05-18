# Outlook 2016 Add-in Bug Demo

In Outlook 2016 15.18 on Mac, it appears that local storage is removed every time the plugin is opened.  This makes it very difficult to persist user data like tokens between sessions.

### Setup

To run the demo, first `npm install && bower install` to install dependencies.
Run `gulp serve-static` to start the server at localhost:8443.
Install the add-in manifest in the project to Outlook.

### Behavior

On each visit, the add-in will check local storage for the key foo and if there is a value, it will display that value and then remove it from LS.  If there is not a value, the value bar will be added for foo.

On Outlook 2016 15.22 on Mac, it removes local storage values after some amount of time (~15 min), so you'll only see `no foo`.

Steps to repro:
1. Open the add-in in Outlook 2016 for Mac 15.22.
2. Close the add-in.
3. Open the add-in again to verify it shows `bar`.
4. Close the add-in again.
5. Wait ~15 min
6. Open the add-in and it should display `no foo`.
