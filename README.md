# word-macros

## RaUI

GUI implementated version of some of the macros in RaMacros

## RaMacros

Word macros to ease some common chores and tasks.

### RangeIsField

- **Type:** Function

- Returns true if rgArg is part of a field

### RangeGetCompleteOutlineLevel

- **Type:** Function

- Returns the complete range from the vPara outline level

### RangeStoryExist

- **Type:** Function

- Returns true if a story with iStory index exist in dcArg document

### StylesDeleteUnused

- **Type:** Function

- Deletes unused styles using multiple loops to respect their hierarchy (avoiding the deletion of fathers without use, like in lists)

> **Based on:** <https://word.tips.net/T001337_Removing_Unused_Styles.html>
>
> **Modifications:**
>
> - Renamed variables
> - It runs until no unused styles left
> - A message with the number of styles must be turn on by the bMsgBox parameter
> - It's now a function that returns the number of deleted styles
> - If the style cannot be found because the NameLocal property is corrupted (eg. because of leading or trailing spaces) it gets automatically deleted
> - It now detects textframes in shapes or inline shapes

### StyleInUse

- **Type:** Function

> Same developer of StylesDeleteUnused

- Is Stryname used any of dcArg's story

### StyleInUseInRangeText

- **Type:** Function

> Same developer of StylesDeleteUnused

- Returns True if "stStyName" is use in rgArg

### StoryInUse

- **Type:** Function

> Same developer of StylesDeleteUnused

- **Note:** this will mark even the always-existing stories as not in use if they're empty

### StyleExists

- **Type:** Function

- Checks if styObjective exists in dcArg and returns a boolean

### StylesDirectFormattingReplace

- **Type:** Subroutine

- Converts bold and italic direct style formatting into Strong and Emphasis

#### Params

- **rgArg:** if nothing the sub works over all the story ranges
- **styUnderline:** the underlined text gets this style applied. It supersedes iUnderlineSelected
- **iUnderlineSelected:** the wdUnderline to be deleted/replaced. It cannot be 0 (wdUnderlineNone)
  - **-1:** default deletes all underline styles
  - **-2:** no underline styles are changed

### StyleSubstitution

- **Type:** Function

- Replace one style with another over the entire document.

#### Params

- **vStyOriginal:** name of the style to be substituted
- **vStySubstitute:** substitute style
- **bDelete:** if True, vStyOriginal will be deleted

#### Returns

- **0:** all good
- **1:** vStyOriginal doesn't exist
- **2:** vStySubstitute doesn't exist
- **3:** neither vStyOriginal nor vStySubstitute exists
- **4:** vStyOriginal and vStySubstitute are the same

### FieldsUnlink

- **Type:** Subroutine

- Unlinks included and embed fields so the images doesn't corrupt the file when it
- gets updated from older (or different software) versions

### FileCopy

- **Type:** Subroutine

- Copies dcArg adding the suffix and/or prefix passed as arguments. In case
- there are none, it appends a number

### FileGetExtension

- **Type:** Function

- Returns the extension of the file

### FileGetNameWithoutExt

- **Type:** Function

- Returns the name of the file without its extension

### FileSaveAsNew

- **Type:** Function

- Saves a copy of the range or document passed as an argument, maintaining the original one opened

#### Params

- **stNewName:** the new document's name
- **stPrefix:** string to prefix the new document's name
- **stSuffix:** string to suffix the new document's name. By default it will add the current date
- **stPath:** the new document's path. If empty it will copy the original's one
- **bOpen:** if True the new document stays open, if false it's saved AND closed
- **bCompatibility:** if True the new document will be converted to the new Word Format
- **bVisible:** if false the new document will be invisible

### FindResetProperties

- **Type:** Subroutine

- Resets find object of rgArg or Selection, if rgArg is Nothing

### FormatNoHighlight

- **Type:** Subroutine

- Takes all highlighting of the document off

#### Params

- **dcArg:** target document

### FormatNoShading

- **Type:** Subroutine

- Takes all shading out of the selected range (or all document if rgArg is Nothing)

#### Params

- **rgArg:** target range. If nothing the sub will loop through the main storyranges
- **dcArg:** if rgArg is Nothing the sub will loop through the storyranges of dcArg

### HeadersFootersRemove

- **Type:** Subroutine

- Removes all headers and footers

### ListsNoExtraNumeration

- **Type:** Subroutine

- Deletes lists' manual numerations

### CleanBasic

- **Type:** Subroutine

- CleanSpaces + CleanEmptyParagraphs
- It's important to execute the subroutines in the proper order to achieve their optimal effects

#### Params

- **rgArg:** the range that will be cleaned. If Nothing it will iterate over all the storyranges of the document
- **bTabs:** if True Tabs are substituted for a single space
- **bBreakLines:** manual break lines get converted to paragraph marks
- **dcArg:** the target document. Necessary in case rgArg is Nothing

### CleanSpaces

- **Type:** Subroutine

- **Deletes:**
  - Tabulations
  - More than 1 consecutive spaces
  - Spaces just before paragraph marks, stops, parenthesis, etc.
  - Spaces just after paragraph marks

#### Params

- **rgArg:** the range that will be cleaned. If Nothing it will iterate over all the storyranges of the document
- **bTabs:** if True Tabs are substituted for a single space
- **dcArg:** the target document. Necessary in case rgArg is Nothing

### CleanEmptyParagraphs

- **Type:** Subroutine

- Deletes empty paragraphs

#### Params

- **rgArg:** the range that will be cleaned. If Nothing it will iterate over all the storyranges of the document
- **bBreakLines:** manual break lines get converted to paragraph marks
- **dcArg:** the target document. Necessary in case rgArg is Nothing

### HeadingsNoPunctuation

- **Type:** Subroutine

- Elimina los puntos finales de los títulos

### HeadingsNoNumeration

- **Type:** Subroutine

- Deletes headings' manual numerations

### HeadingsChangeCase

- **Type:** Subroutine

- Changes the case for the heading selected. This subroutine transforms the text, it doesn't change the style option "All caps"

#### Params

- **dcArg:** the document to be changed
- **iHeading:** the heading style to be changed. If 0 all headings will be processed
- **iCase:** the desired case for the text. It can be one of the WdCharacterCase constants. Options:
  - **wdLowerCase:** 0
  - **wdUpperCase:** 1
  - **wdTitleWord:** 2
  - **wdTitleSentence:** 4
  - **wdToggleCase:** 5

### HyperlinksDeleteAll

- **Type:** Subroutine

- Deletes all hyperlinks

### HyperlinksFormatting

- **Type:** Subroutine

- It cleans and format hyperlinks

#### Params

- **iPurpose:** choose what is the aim of the subroutine:
  - **1:** Applies the hyperlink style to all hyperlinks
  - **2:** cleans the text showed so only the domain is left
  - **3:** both

### ImagesToCenteredInLine

- **Type:** Subroutine

- Image formatting
  - From shapes to inlineshapes
  - Centered
  - Correct aspect ratio
  - No bigger than the page

### QuotesStraightToCurly

- **Type:** Subroutine

- Changes problematic straight quote marks (" and ') to curly quotes
- Deletes the non configurable variables of Document.Autoformat

### SectionBreakBeforeHeading

- **Type:** Subroutine

- Inserts section breaks of the type assigned before each heading of the level selected

#### Params

- **dcArg:** the document to be changed
- **bRespect:** respect the original section start type before the heading
- **iWdSectionStart:** the kind of section break to insert
- **iHeading:** heading style that will be found

### SectionGetFirstFootnoteNumber

- **Type:** Function

- Returns the number of the first footnote of the section or 0 if there is none

#### Params

- **lIndex:** the index of the section containing the footnote

### SectionsExportEachToFiles

- **Type:** Function

- Exports each section of the document to a separate file

#### Params

- **bClose:** if true, close the section documents after exporting them
- **bMaintainFootnotesNumeration:** if true, maintain the same footnote numeration for each section
- **bMaintainPagesNumeration:** if true, maintain the same page numeration for each section
- **stNewDocName:** name of the exported file
- **stPrefix:** prefix for stNewDocName
- **stSuffix:** suffix for stNewDocName
- **stPath:** path of the exported file

### SectionsFillBlankPages

- **Type:** Subroutine

- Puts a blank page before each even or odd section break

#### Params

- **dcArg:** the document to be changed
- **stFillerText:** an optional dummy string to fill the blank page
- **styFillStyle:** style for the dummy text

### TablesConvertToImage

- **Type:** Subroutine

- Convert each table to an inline image
- **Solution to problems with clipboard (do loop) found in [Mr Excel](https:**//www.mrexcel.com/board/threads/excel-vba-inconsistent-errors-when-trying-to-copy-and-paste-objects-from-excel-to-word.1112368/post-5485704)

#### Params

- **iPlacement:** WdOLEPlacement enum
  - **0:** wdInLine
  - **1:** wdFloatOverText

### TablesConvertToText

- **Type:** Subroutine

- Convert each table in the range to text
- If no range is passed as an argument, it will act on the selection

#### Params

- **iSeparator:** the column separator parameter:
  - wdSeparateByParagraphs 			0
  - wdSeparateByTabs 					1
  - wdSeparateByCommas 				2
  - wdSeparateByDefaultListSeparator 	3
- **bNested:** the NestedTables parameter

### TablesExportToNewFile

- **Type:** Function

- Export each table of the selected range to a new document

#### Params

- **rgArg:** if nothing the tables in the Content range of dcArg will be exported
- **dcArg:** it will get supersede by the parent of rgArg if it isn't nothing
- **bSameMarkUp:** if true the new document will be a copy of the current one, but blank
- **vTemplate:** if bSameMarkUp is false the new document will be based on vTemplate
- **stDocName:** name of the parent document
- **stDocPrefix:** the prefix to append to the new document
- **stDocSuffix:** the suffix to append to the new document
- **stPath:** the new document's path
- **iBreak:** WdBreakType that will follow each table. (If 0 there won't be any. Default: 7 (wdPageBreak))
- **bTitles:** if true a text will precede each table with it's own text
- **stTitle:** the text to insert if the table doesn't have a title
- **bOverwrite :** if true the table titles will be replaced by stTitle
- **vTitleStyle:** the style of the headings

### TablesExportToPdf

- **Type:** Subroutine

- Export each table of the argument range to a PDF file

#### Params

- **stPath:** path of the documents
- **stDocName:** name of the parent document
- **stTableSuffix:** the suffix to append to the table title, if it hasn't any
- **bDelete:** defines if the table should be replaced
- **stReplacementText:** the replacement text before the table title
- **bLink:** if true the replacement text will be a hyperlink pointing to the address of the pdf
- **stAddress:** the path where the hyperlink will point.
  - The name of the file will be automatically added to the argument, BUT
  - If empty it will point to the destination of the exported pdf
- **vStyle:** the paragraph style of the replacement text
- **iSize:** the font size of the replacement text
- **bFullPage:** if true the table is exported along with the rest of its page
- **bExport:** if false the table will be processed but not exported

### TablesStyle

- **Type:** Subroutine

- Formats all tables within rgArg or dcArg with vStyle

### FootnotesDeleteEmpty

- **Type:** Subroutine

- Deletes empty footnotes that had been incorrectly manually erased

### FootnotesFormatting

- **Type:** Subroutine

- Applies styles to the footnotes story and the footnotes references

#### Params

- **stFootnotes:** style for the body text. Default: wdStyleFootnoteText
- **styFootnoteReferences:** style for the references. Default: stFootnoteReferences

### FootnotesHangingIndentation

- **Type:** Subroutine

- Adds a tab to the beginning of each paragraph and footnote, so their indentation is hanging

#### Params

	- **sIndentation:** the position of the indented text in centimeters
	- **iFootnoteStyle:** to indicate a custom footnote style

### FootnotesSameNumberingRule

- **Type:** Subroutine

- Set the same footnotes numbering rule in all sections of the document

#### Params

- **iNumberingRule:**
  - **3 (default):** it gives the numbering rule of the first section to all others
  - **0:** wdRestartContinuous
  - **1:** wdRestartSection
  - **2:** wdRestartPage: wdRestartPage
- **iStartingNumber:** starting number of each section
  - **-501:** doesn't change anything
  - **0:** copies the starting number of the first section in all the others

### ClearHiddenText

- **Type:** Function

- Deletes or apply a warning style to all hidden text in the document.

#### Returns

- Array of integers of the story ranges containing hidden text

#### Params

- **bDelete:** true deletes all hidden text
- **styWarning:** defines the style for the hidden text
- **bMaintainHidden:** if true the text maintains its hidden attribute
- **bShowHidden:** changes if the hidden text is displayed
  - **0:** maintains the current configuration
  - **1:** hidden
  - **2:** visible
