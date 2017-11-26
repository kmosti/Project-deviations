## prosjekt-avvik

The web part is intended to be used with a project site that contains a project info list and a deviation list on a different site collection.

In a next version, these settings will be configurable from the settings pane, but right now they are hard coded.

The list will display deviations for the current project (current site), and will have a button for adding new deviations, which is handled by a PowerApp.

There is also a current minor issue on mobile view where the "add deviation" button is set to display:none.

### Preview

![Demo](https://user-images.githubusercontent.com/20144749/33244403-cb0f2fa2-d2f6-11e7-96a7-c9a19043a845.gif)

### Building the code

```bash
git clone the repo
npm i
npm i -g gulp
gulp
```

This package produces the following:

* lib/* - intermediate-stage commonjs build artifacts
* dist/* - the bundled script, along with other resources
* deploy/* - all resources which should be uploaded to a CDN.
