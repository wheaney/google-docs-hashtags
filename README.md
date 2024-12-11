# google-docs-hashtags

Google Docs Apps Script for indexing hashtags. Intended for journaling, allowing for tagging people and other categories that are important to you. Tag indexing will allow you to easily view all entries that use that tag.

## Document setup

The script makes a couple assumptions about the Google Doc it's being used with:
1. Tagged content must be within sections that begin with `Heading 3` styled text. The code still calls these section headings "dates" because it was written for a journal, where tags are indexed from within dated journal entries. When viewing content in the tags index, it will link back to these sections using the heading text.
2. The document must have a `Heading 1` whose text *exactly matches* "Tags". **Everything after this heading is maintained by the script and subject to deletion** and everything before this heading is subject to indexing. In general, this heading should go at the end of the document if you've never run the script before.
3. Tags are expected to be hashtags, prefixed with `#`, typically at the end of a sentence. The content shown in the index will be just the paragraph/line that contains the tag. Some notes about hashtag usage:
* Multiple tags are allowed per line
* If you want a tag to encompass more than one paragraph/line, you can add `_+d` to the end of the hashtag where `d` is replaced by the number of lines following this that should be indexed.
* Images and lists will also be indexed, but since lists are multi-line, be sure to follow the "more than one line" guidance above

### Journal setup

If used specifically for journaling:
* Use `Heading 1` for years
* Use `Heading 2` for months
* Use `Heading 3` for days (I put the full date string here)

## Script setup

1. Open the Google Doc you want to use this with. 
2. Go to `Extensions` / `Apps Script`. 
3. From the `Code.gs` tab, copy and paste the contents of `AppsScript.gs` into the editor and save it. 
4. From the `functions` dropdown next to the Run/Debug buttons above the editor, choose `findTagsAndBuildIndex`.
5. Once you've set up the document with the appropriate headings as described in the last section, click `Run` and leave this browser tab open until it completes. 
6. When it's done, the `Tags` section of the document will be filled out.