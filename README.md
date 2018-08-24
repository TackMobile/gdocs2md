gdocs2md
========

A simple Google Apps script to convert a properly formatted Google Drive Document to the markdown (.md) format. 

## Usage

  * Adding this script to your doc (once per doc):
    * Open your Google Drive document (http://drive.google.com)
    * Tools -> Script Manager > New
    * Select "Blank Project", then paste this code in and save.
    * Clear the myFunction() default empty function and paste the contents of [converttomarkdown.gapps](https://raw.githubusercontent.com/TackMobile/gdocs2md/master/converttomarkdown.js) into the code editor
    * File -> Save
    
  * Running the script (run as many times as you want):
    - Tools > Script Manager
    - Select "ConvertToMarkdown" function.
    - Click Run button (First run will require you to authorize it. Authorize and run again)
    - Converted doc with images attached will be emailed to you. Subject will be "[MARKDOWN_MAKER]...".


## Interpreted formats
  * Text:
    * paragraphs are separated by two newlines
    * text styled as heading 1, 2, 3, etc is converted to Markdown heading: #, ##, ###, etc
    * text formatted with Courier New is backquoted: ``text``
    * links are converted to MD format: `[anchortext](url)`
  * Lists:
    * Numbered lists are converted correctly, including nested lists
    * bullet lists are converted to "`*`" Markdown format appropriately, including nested lists
  * Images:
    * images are correctly extracted and sent as attachments
  * Blocks:
    * Table of contents is replaced by `[[TOC]]`
    * blocks of text delimited by "--- class whateverclassnameyouwant" and "---" are converted to `<div class="whateverclassnameyouwant"></div>` 
    * Source code: 
      * **UPDATED**: blocks of text delimited by "--- source code" or "--- src" and "---" are converted to `<pre></pre>`
      * **NEW**: blocks of text delimited by "--- source pretty" or "--- srcp" and "---" are converted to `<pre class="prettyprint"></pre>`
      * Source blocks delimited with "\`\`\`srcType" (eg: \`\`\`java) and ``` are respected and kept intact
    * Jekyll Header:
      * Block delimited with "---header" and "---" should contain at least the title in format "title: "Some Title"" and author in format "author: Author Name". A "layout: post" and "date: 2017-09-27" will be appended to the header. If images are pressent, a blogImageDir value should be present "blogImageDir: directoryImagesWillBeIn"
    * Tables:
      * **NEW**: Simple `<table>` processing
  * "--- jsperf `<testID>`" is replaced by an iframe that shows an interactive chart of a JSPerf test. The `<testID>` is the last part of the URL of the Browserscope anchor in your JSPerf test. Something like `"agt1YS1wcm9maWxlcnINCxIEVGVzdBjlm_EQDA"` in the URL `http://www.browserscope.org/user/tests/table/agt1YS1wcm9maWxlcnINCxIEVGVzdBjlm_EQDA`
 


## CONTRIBUTORS

* Renato Mangini - [G+](//google.com/+renatomangini) - [Github](//github.com/mangini)
* Ed Bacher - [G+](//plus.google.com/106923847899206957842) - [Github](//github.com/evbacher)

## LICENSE

Use this script at your will, on any document you want and for any purpose, commercial or not. 
The MarkDown files generated by this script are not considered derivative work and 
don't require any attribution to the owners of this script. 

If you want to modify and redistribute the script (not the converted documents - those are yours), 
just keep a reference to this repo or to the license info below:

```
Copyright 2013 Google Inc. All Rights Reserved.

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
```
