# Release notes: illiad-addon-blacklight


## [1.6.1] - 2026-09-03

### Changed
- Fix the Tab keyboard navigation bug - a known issue with the WebView2 browser, this fix will let users tab between the Detail, History, ..., Blacklight OPAC Search tabs without hijacking the focus
- Moved the JournalInfo to the top, above the browser

### Added
- Newly generated layout xml file created from the windows client using the customization tool found when right clicking on the addon pane

## [1.6] - 2026-09-02

### Changed
- To support the bootstrap 5 changes in catalog, the addon needs to use a modern browser type (WebView2)

### Added
- Release notes file

### Removed
- No longer need the layout xml files, they were used for the old browser
- Removed the TOU Info pane, not used
