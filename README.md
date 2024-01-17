# DOCX.MERGER
### Video Demo:  https://youtu.be/KPIsSLJ57IQ
### Description:
#### Summary

**Released:** October 21, 2023  
**Version:** 1.0  
**Written on Python ver.:** 3.12.0  
**Made by:** Vojtěch Ettler

This is a little gadget allowing users to combine two or more Word documents into one while preserving the style and format of the first file selected. The program lets you navigate to a folder via simple custom commands, printing out list of directories at each step. Once you locate the folder with the desired content, the **docx.Merger** shows you all the documents to be combined, automatically skipping any unrelated files, and asks you one last time for confirmation. As soon as you give it, the merged file is created (it may take some time depending on the number and length of the docs), and the program outputs its word count, character count and number of standard pages. Quick and easy for anyone who works with MS Word on a regular basis.

#### Technical details

The program is written in the Python programming language and runs on several imported modules. Among them the most important for the app's functioning are `python-docx` and `docxcompose` libraries which make it possible to read and edit Word documents. Big thank you to their authors!    For using this program, it is necessary to install these modules beforehand (see *requirements.txt*). 
  
There are three core function in the code apart from `main`: `get_path` substituting common file browsing, `merger` taking care of the putting files together, and `text_counter` summing up text statistics of the result. The `get_path` function heavily depends on class `Path` introduced to organize the code in a more transparent way thanks to several methods that manage the user input of keywords.
The code is fully commented and counts 130 lines.
 
#### Testing

As for automated testing, there is *test_project.py* included in the root folder. It checks all  three principal functions, mocking the class instance and methods for the purposes of examining the `get_path` function. All tests pass successfully, be it through the `pytest` or .py file itself.

Testing is done on *test1.docx* and *test2.docx* in *test* folder (the folder contains a dummy .py file to verify that the merger is able to skip unsuitable items).

One deprecation warning is currently listed by pytest due to `docxcompose` module using an obsolete "import pkg_resources" as API.


#### *Possible improvements*

* creating GUI which will make for much friendlier user experience
* bundling the program into executable via `pyinstaller`
* rewriting it in a way so the `DeprecationWarning` is not raised

#### Background info

And now for something more personal. What inspired me to create such a gimmick? Being a literary translator (for the time being)! Every time I finish my work on a book, I need to combine up to ten Word documents containing parts of my translation, and doing it by in-built Office 365 option or manually through copy and paste is inconvenient. There you have it: Python to the rescue!

### <p style="text-align: center;"> Happy merging. </p>


