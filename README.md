# word-macros

> At the moment I'm not using Word anymore, so this project won't see any further development

Macros developed to ease my work with Microsoft Word.

The functions are in RaMacros.bas and RaUI.bas contains some GUI implementations.

## RangeIsField

**Returns** true if rgArg is part of a field.

## RangeGetCompleteOutlineLevel

**Returns** the complete range from the vPara outline level.

## RangeStoryExist

**Returns** true if a story with iStory index exist in dcArg document.

## StylesDeleteUnused

Deletes unused styles using multiple loops to respect their hierarchy (avoiding the deletion of fathers without use, like in lists).

> **Based on:** <https://word.tips.net/T001337_Removing_Unused_Styles.html>.

> **Modifications:**
>
> - Renamed variables.
> - It runs until no unused styles left.
> - A message with the number of styles must be turn on by the bMsgBox parameter.
> - It's now a function that **Returns** the number of deleted styles.
> - If the style cannot be found because the NameLocal property is corrupted (eg. because of leading or trailing spaces) it gets automatically deleted.
> - It now detects textframes in shapes or inline shapes.

## StyleInUse

> **Based on:** <https://word.tips.net/T001337_Removing_Unused_Styles.html>.

**Returns** a boolean after checking if *stStyName* is used in any of *dcArg's* stories.

### Params

1. **stStyName:** style to check.
2. **dcArg:** target document.

## StyleInUseInRangeText

> **Based on:** <https://word.tips.net/T001337_Removing_Unused_Styles.html>.

**Returns** True if "stStyName" is use in rgArg.

## StoryInUse

> **Based on:** <https://word.tips.net/T001337_Removing_Unused_Styles.html>.

> **Note:** this will mark even the always-existing stories as not in use if they're empty.

## StyleExists

**Returns** a boolean after checking if styObjective exists in dcArg.

## StylesDirectFormattingReplace

Converts bold and italic direct style formatting into Strong and Emphasis.

### Params

1. **rgArg:** if nothing the sub works over all the story ranges.
1. **styUnderline:** the underlined text gets this style applied. It supersedes iUnderlineSelected.
1. **iUnderlineSelected:** the wdUnderline to be deleted/replaced. It cannot be 0 (wdUnderlineNone).
    - **-1:** default deletes all underline styles
    - **-2:** no underline styles are changed

## StyleSubstitution

Replace one style with another over the entire document.

### Params

1. **vStyOriginal:** name of the style to be substituted.
1. **vStySubstitute:** substitute style.
1. **bDelete:** if True, vStyOriginal will be deleted.

### Returns

- **0:** all good.
- **1:** vStyOriginal doesn't exist.
- **2:** vStySubstitute doesn't exist.
- **3:** neither vStyOriginal nor vStySubstitute exists.
- **4:** vStyOriginal and vStySubstitute are the same.

## FieldsUnlink

Unlinks *included* and *embedded* fields so the images doesn't corrupt the file when it gets updated from older (or different software) versions.

## FileCopy

Copies dcArg adding the suffix and/or prefix passed as arguments. In case there are none, it appends a number.

## FileGetExtension

**Returns** the extension of the file.

## FileGetNameWithoutExt

**Returns** the name of the file without its extension.

## FileSaveAsNew

Saves a copy of the range or document passed as an argument, maintaining the original one opened.

### Params

1. **stNewName:** the new document's name.
1. **stPrefix:** string to prefix the new document's name.
1. **stSuffix:** string to suffix the new document's name. By default it will add the current date.
1. **stPath:** the new document's path. If empty it will copy the original's one.
1. **bOpen:** if True the new document stays open, if false it's saved AND closed.
1. **bCompatibility:** if True the new document will be converted to the new Word Format.
1. **bVisible:** if false the new document will be invisible.

## FindResetProperties

Resets find object of rgArg or Selection, if rgArg is Nothing.

## FormatNoHighlight

Takes all highlighting of the document off.

### Params

1. **dcArg:** target document.

## FormatNoShading

Takes all shading out of the selected range (or all document if rgArg is Nothing).

### Params

1. **rgArg:** target range. If nothing the sub will loop through the main storyranges.
1. **dcArg:** if rgArg is Nothing the sub will loop through the storyranges of dcArg.

## HeadersFootersRemove

Removes all headers and footers.

## ListsNoExtraNumeration

Deletes lists' manual numerations.

## CleanBasic

> CleanSpaces + CleanEmptyParagraphs.

It's important to execute the subroutines in the proper order to achieve their optimal effects.

### Params

1. **rgArg:** the range that will be cleaned. If Nothing it will iterate over all the storyranges of the document.
1. **bTabs:** if True Tabs are substituted for a single space.
1. **bBreakLines:** manual break lines get converted to paragraph marks.
1. **dcArg:** the target document. Necessary in case rgArg is Nothing.

## CleanSpaces

**Deletes:**

- Tabulations.
- More than 1 consecutive spaces.
- Spaces just before paragraph marks, stops, parenthesis, etc.
- Spaces just after paragraph marks.

### Params

1. **rgArg:** the range that will be cleaned. If Nothing it will iterate over all the storyranges of the document.
1. **bTabs:** if True Tabs are substituted for a single space.
1. **dcArg:** the target document. Necessary in case rgArg is Nothing.

## CleanEmptyParagraphs

Deletes empty paragraphs.

### Params

1. **rgArg:** the range that will be cleaned. If Nothing it will iterate over all the storyranges of the document.
1. **bBreakLines:** manual break lines get converted to paragraph marks.
1. **dcArg:** the target document. Necessary in case rgArg is Nothing.

## HeadingsNoPunctuation

Elimina los puntos finales de los títulos.

## HeadingsNoNumeration

Deletes headings' manual numerations.

## HeadingsChangeCase

Changes the case for the heading selected. This subroutine transforms the text, it doesn't change the style option "All caps".

### Params

1. **dcArg:** the document to be changed.
1. **iHeading:** the heading style to be changed. If 0 all headings will be processed.
1. **iCase:** the desired case for the text. It can be one of the WdCharacterCase constants.  .
Options:
    - **0:** wdLowerCase.
    - **1:** wdUpperCase.
    - **2:** wdTitleWord.
    - **4:** wdTitleSentence.
    - **5:** wdToggleCase.

## HyperlinksDeleteAll

Deletes all hyperlinks.

## HyperlinksFormatting

It cleans and format hyperlinks.

### Params

1. **iPurpose:** choose what is the aim of the subroutine:
    - **1:** Applies the hyperlink style to all hyperlinks.
    - **2:** cleans the text showed so only the domain is left.
    - **3:** both.

## ImagesToCenteredInLine

Image formatting:

- From shapes to inlineshapes.
- Centered.
- Correct aspect ratio.
- No bigger than the page.

## QuotesStraightToCurly

Changes problematic straight quote marks (" and ') to curly quotes and deletes the non configurable variables of Document.Autoformat.

## SectionBreakBeforeHeading

Inserts section breaks of the type assigned before each heading of the level selected.

### Params

1. **dcArg:** the document to be changed.
1. **bRespect:** respect the original section start type before the heading.
1. **iWdSectionStart:** the kind of section break to insert.
1. **iHeading:** heading style that will be found.

## SectionGetFirstFootnoteNumber

**Returns** the number of the first footnote of the section or 0 if there is none.

### Params

1. **lIndex:** the index of the section containing the footnote.

## SectionsExportEachToFiles

Exports each section of the document to a separate file.

### Params

1. **bClose:** if true, close the section documents after exporting them.
1. **bMaintainFootnotesNumeration:** if true, maintain the same footnote numeration for each section.
1. **bMaintainPagesNumeration:** if true, maintain the same page numeration for each section.
1. **stNewDocName:** name of the exported file.
1. **stPrefix:** prefix for stNewDocName.
1. **stSuffix:** suffix for stNewDocName.
1. **stPath:** path of the exported file.

## SectionsFillBlankPages

Puts a blank page before each even or odd section break.

### Params

1. **dcArg:** the document to be changed.
1. **stFillerText:** an optional dummy string to fill the blank page.
1. **styFillStyle:** style for the dummy text.

## TablesConvertToImage

Convert each table to an inline image.

> Solution to problems with clipboard (do loop) found in [Mr Excel](https:**//www.mrexcel.com/board/threads/excel-vba-inconsistent-errors-when-trying-to-copy-and-paste-objects-from-excel-to-word.1112368/post-5485704).

### Params

1. **iPlacement:** WdOLEPlacement enum.
    - **0:** wdInLine.
    - **1:** wdFloatOverText.

## TablesConvertToText

Converts each table in the range to text. If no range is passed as an argument, it will act on the selection.

### Params

1. **iSeparator:** the column separator parameter.
    - **0:** wdSeparateByParagraphs.
    - **1:** wdSeparateByTabs.
    - **2:** wdSeparateByCommas.
    - **3:** wdSeparateByDefaultListSeparator.
2. **bNested:** the NestedTables parameter.

## TablesExportToNewFile

Export each table of the selected range to a new document.

### Params

1. **rgArg:** if nothing the tables in the Content range of dcArg will be exported.
1. **dcArg:** it will get supersede by the parent of rgArg if it isn't nothing.
1. **bSameMarkUp:** if true the new document will be a copy of the current one, but blank.
1. **vTemplate:** if bSameMarkUp is false the new document will be based on vTemplate.
1. **stDocName:** name of the parent document.
1. **stDocPrefix:** the prefix to append to the new document.
1. **stDocSuffix:** the suffix to append to the new document.
1. **stPath:** the new document's path.
1. **iBreak:** WdBreakType that will follow each table. (If 0 there won't be any. Default: 7 (wdPageBreak)).
1. **bTitles:** if true a text will precede each table with it's own text.
1. **stTitle:** the text to insert if the table doesn't have a title.
1. **bOverwrite :** if true the table titles will be replaced by stTitle.
1. **vTitleStyle:** the style of the headings.

## TablesExportToPdf

Export each table of the argument range to a PDF file.

### Params

1. **stPath:** path of the documents.
1. **stDocName:** name of the parent document.
1. **stTableSuffix:** the suffix to append to the table title, if it hasn't any.
1. **bDelete:** defines if the table should be replaced.
1. **stReplacementText:** the replacement text before the table title.
1. **bLink:** if true the replacement text will be a hyperlink pointing to the address of the pdf.
1. **stAddress:** the path where the hyperlink will point.  .
The name of the file will be automatically added to the argument, **BUT** if empty it will point to the destination of the exported pdf.
1. **vStyle:** the paragraph style of the replacement text.
1. **iSize:** the font size of the replacement text.
1. **bFullPage:** if true the table is exported along with the rest of its page.
1. **bExport:** if false the table will be processed but not exported.

## TablesStyle

Formats all tables within rgArg or dcArg with vStyle.

## FootnotesDeleteEmpty

Deletes empty footnotes that had been incorrectly manually erased.

## FootnotesFormatting

Applies styles to the footnotes story and the footnotes references.

### Params

1. **stFootnotes:** style for the body text. Default: wdStyleFootnoteText.
1. **styFootnoteReferences:** style for the references. Default: stFootnoteReferences.

## FootnotesHangingIndentation

Adds a tab to the beginning of each paragraph and footnote, so their indentation is hanging.

### Params

1. **sIndentation:** the position of the indented text in centimeters.
1. **iFootnoteStyle:** to indicate a custom footnote style.

## FootnotesSameNumberingRule

Set the same footnotes numbering rule in all sections of the document.

### Params

1. **iNumberingRule:**
    - **3 (default):** it gives the numbering rule of the first section to all others.
    - **0:** wdRestartContinuous.
    - **1:** wdRestartSection.
    - **2:** wdRestartPage: wdRestartPage.
1. **iStartingNumber:** starting number of each section.
    - **-501:** doesn't change anything.
    - **0:** copies the starting number of the first section in all the others.

## ClearHiddenText

Deletes or apply a warning style to all hidden text in the document.

**Returns** an array of integers of the story ranges containing hidden text.

### Params

1. **bDelete:** true deletes all hidden text.
1. **styWarning:** defines the style for the hidden text.
1. **bMaintainHidden:** if true the text maintains its hidden attribute.
1. **bShowHidden:** changes if the hidden text is displayed.
    - **0:** maintains the current configuration.
    - **1:** hidden.
    - **2:** visible.
