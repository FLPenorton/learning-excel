
# Table of Contents

1.  [Getting Started](#org0e56f14)
    1.  [Ribbon Menu and Quick Access Toolbar](#org18f59be)
        1.  [Home](#org831afbc)
        2.  [Insert - Add additional objects](#org481e7b8)
        3.  [Page Layout](#org188cf24)
        4.  [Formulas](#org722158a)
        5.  [Data](#orgd3842d8)
        6.  [Review](#org3d7b4cb)
        7.  [Help](#org2b6695c)
    2.  [Shortcut menus and mini toolbar](#orgf20c841)
        1.  [Cells and Texts in Cells](#org08ed975)
        2.  [Workbook and Worksheets](#org7f27ea3)
    3.  [Getting Help](#org8aa4947)
    4.  [Accessibility](#org67503fb)
2.  [Data Entry](#orgeee5581)
    1.  [Entry and Autofill](#org2b8f5fd)
    2.  [Dates and Times](#org77945ec)
    3.  [Undo and Redo](#org89caed9)
    4.  [Save or Save As](#org18e8935)
3.  [Formulas and Functions](#org745d1fa)
    1.  [Simple Formulas](#org578fb7d)
    2.  [Copying formulas](#orgcb57a9c)
    3.  [Modifying formulas](#org56c7386)
    4.  [Using SUM, AVERAGE, COUNT, COUNTA](#org0967df0)
    5.  [XLOOKUP, and other lookup functions](#org195f465)
4.  [Formatting](#orge18e5fc)
    1.  [Font styles, borders, and background colors](#orgf87f480)
    2.  [Formatting numbers](#org9a7264f)
    3.  [Worksheets](#org7149f13)
5.  [Preparing to print](#org7097dc7)
    1.  [Page Layout](#org401a4df)
    2.  [Print Setup and Page Break Preview](#org0114021)
6.  [Charts](#org500b16c)
    1.  [Creating](#orgdd104c6)
    2.  [Chart Types](#org082d8bd)
7.  [Worksheet Views](#orgd4875d5)
    1.  [Freezing and Unfreezing Panes](#orga05daa1)
    2.  [Split Screens](#orgffea8ca)
    3.  [Using Worksheets and Workbooks](#orgc0530fe)
        1.  [Insert, Delete, and Rename](#orga749989)
        2.  [Move, Copy, and Group](#org1523267)
        3.  [Open, Close, and Save](#org65f4bc6)
8.  [Data Management](#org230e6db)
    1.  [Sort](#org4bbd684)
    2.  [Filter](#orgca8ff1e)
    3.  [Make a Pivot Table](#org3bdd367)
9.  [Security and Sharing](#orge941d3b)
    1.  [Protect worksheets and workbooks](#org7864bd5)
    2.  [Share and Track changes](#orgb17f91b)
10. [Common Actions Across Applications and Windows OS](#org31392c3)
11. [Modifiers for Movement and Selection](#org53c35d7)



<a id="org0e56f14"></a>

# Getting Started


<a id="org18f59be"></a>

## Ribbon Menu and Quick Access Toolbar


<a id="org831afbc"></a>

### Home

Shows most commonly use actions
Clipboard
Font
Text Alignment
Number Formatting
Cell Styles
Cell Manipulation
Cell Editing


<a id="org481e7b8"></a>

### Insert - Add additional objects

Tables
Illustrations
Charts


<a id="org188cf24"></a>

### Page Layout


<a id="org722158a"></a>

### Formulas


<a id="orgd3842d8"></a>

### Data


<a id="org3d7b4cb"></a>

### Review


<a id="org2b6695c"></a>

### Help


<a id="orgf20c841"></a>

## Shortcut menus and mini toolbar


<a id="org08ed975"></a>

### Cells and Texts in Cells

Use ALT + ENTER to create a
line break in a cell


<a id="org7f27ea3"></a>

### Workbook and Worksheets

1.  Rows and Columns

2.  Cell Addresses


<a id="org8aa4947"></a>

## Getting Help

Stop and hover on icons to discover for keyboard shortcuts and more, see (font color)
Use Search (for actions)
Hit F1 or Help tab and search for found action to learn more


<a id="org67503fb"></a>

## Accessibility

See Review Tab / Check Accessibility


<a id="orgeee5581"></a>

# Data Entry


<a id="org2b8f5fd"></a>

## Entry and Autofill

Modify Using Doubleclick, F2, or Overtype
To Wrap or not to Wrap


<a id="org77945ec"></a>

## Dates and Times

Dates: stored as numbers starting from 1900 so =32 formatted as date looks like &ldquo;2/1/1900&rdquo;


<a id="org89caed9"></a>

## Undo and Redo

CTRL+z
CTRL+y


<a id="org18e8935"></a>

## Save or Save As


<a id="org745d1fa"></a>

# Formulas and Functions


<a id="org578fb7d"></a>

## Simple Formulas

()      :: grouping

-   **-    :** addition and subtraction

/ \*     :: division and multiplication
^ POWER :: exponentiation


<a id="orgcb57a9c"></a>

## Copying formulas


<a id="org56c7386"></a>

## Modifying formulas


<a id="org0967df0"></a>

## Using SUM, AVERAGE, COUNT, COUNTA


<a id="org195f465"></a>

## XLOOKUP, and other lookup functions

<table border="2" cellspacing="0" cellpadding="6" rules="groups" frame="hsides">


<colgroup>
<col  class="org-left" />

<col  class="org-left" />
</colgroup>
<tbody>
<tr>
<td class="org-left"><b>VLOOKUP</b></td>
<td class="org-left"><b>HLOOKUP</b></td>
</tr>

<tr>
<td class="org-left">Return array must be to the right of the lookup array</td>
<td class="org-left">Position doesn&rsquo;t matter</td>
</tr>

<tr>
<td class="org-left">Must wrap lookup errors with IFERROR()</td>
<td class="org-left">Lookup errors are handled with &rsquo;if not found&rsquo;</td>
</tr>

<tr>
<td class="org-left">To find approximate values, the lookup array must be sorted ascending</td>
<td class="org-left">To find approximate values, the lookup array doesn&rsquo;t need to be sorted</td>
</tr>

<tr>
<td class="org-left">Inserting columns between lookup and return array will return corrupted results</td>
<td class="org-left">Lookups are not column position dependent.</td>
</tr>

<tr>
<td class="org-left">Can&rsquo;t lookup horizontally, must use HLOOKUP</td>
<td class="org-left">Can lookup vertically and/or horizontally</td>
</tr>

<tr>
<td class="org-left">Doesn&rsquo;t default to exact match which is usually what is wanted</td>
<td class="org-left">Defaults to exact match</td>
</tr>

<tr>
<td class="org-left">Cannot do bottom-up search</td>
<td class="org-left">Can do bottom-up search</td>
</tr>
</tbody>
</table>


<a id="orge18e5fc"></a>

# Formatting


<a id="orgf87f480"></a>

## Font styles, borders, and background colors


<a id="org9a7264f"></a>

## Formatting numbers


<a id="org7149f13"></a>

## Worksheets

Height and Width adjustments
Insert and Delete
Hide and Unhide
Moving: Copy, Cut, and Paste
Find and Replace data


<a id="org7097dc7"></a>

# Preparing to print


<a id="org401a4df"></a>

## Page Layout


<a id="org0114021"></a>

## Print Setup and Page Break Preview


<a id="org500b16c"></a>

# Charts


<a id="orgdd104c6"></a>

## Creating


<a id="org082d8bd"></a>

## Chart Types


<a id="orgd4875d5"></a>

# Worksheet Views


<a id="orga05daa1"></a>

## Freezing and Unfreezing Panes


<a id="orgffea8ca"></a>

## Split Screens

Horizontal
Vertical


<a id="orgc0530fe"></a>

## Using Worksheets and Workbooks


<a id="orga749989"></a>

### Insert, Delete, and Rename


<a id="org1523267"></a>

### Move, Copy, and Group


<a id="org65f4bc6"></a>

### Open, Close, and Save


<a id="org230e6db"></a>

# Data Management


<a id="org4bbd684"></a>

## Sort


<a id="orgca8ff1e"></a>

## Filter


<a id="org3bdd367"></a>

## Make a Pivot Table


<a id="orge941d3b"></a>

# Security and Sharing


<a id="org7864bd5"></a>

## Protect worksheets and workbooks


<a id="orgb17f91b"></a>

## Share and Track changes


<a id="org31392c3"></a>

# Common Actions Across Applications and Windows OS

<table border="2" cellspacing="0" cellpadding="6" rules="groups" frame="hsides">


<colgroup>
<col  class="org-left" />

<col  class="org-left" />

<col  class="org-left" />
</colgroup>
<tbody>
<tr>
<td class="org-left"><b>Action</b></td>
<td class="org-left"><b>Keyboard Shortcut</b></td>
<td class="org-left"><b>Description</b></td>
</tr>

<tr>
<td class="org-left">Cut</td>
<td class="org-left">CTRL+X</td>
<td class="org-left">Cut &rsquo;out&rsquo; currently selected object and place it in the clipboard</td>
</tr>

<tr>
<td class="org-left">Copy</td>
<td class="org-left">CTRL+C</td>
<td class="org-left">Copy currently selected object</td>
</tr>

<tr>
<td class="org-left">Paste</td>
<td class="org-left">CTRL+V</td>
<td class="org-left">Paste currently copied object</td>
</tr>

<tr>
<td class="org-left">Edit (inside)</td>
<td class="org-left">F2</td>
<td class="org-left">Edit inside selected object</td>
</tr>

<tr>
<td class="org-left">Select All</td>
<td class="org-left">CTRL+A</td>
<td class="org-left">Select all items in current active region</td>
</tr>

<tr>
<td class="org-left">Close Window</td>
<td class="org-left">CTRL+W</td>
<td class="org-left">Close the current application instance window.  Doesn&rsquo;t close the application.</td>
</tr>

<tr>
<td class="org-left">Undo</td>
<td class="org-left">CTRL+Z</td>
<td class="org-left">Undo (reverse) past actions</td>
</tr>

<tr>
<td class="org-left">Redo</td>
<td class="org-left">CTRL+Y</td>
<td class="org-left">Redo (reverse undo)</td>
</tr>

<tr>
<td class="org-left">Save</td>
<td class="org-left">CTRL+S</td>
<td class="org-left">Save changes document, worksheet(s), workbook, etc.</td>
</tr>

<tr>
<td class="org-left">Print</td>
<td class="org-left">CTRL+P</td>
<td class="org-left">Print document, worksheet(s), workbook, etc.</td>
</tr>

<tr>
<td class="org-left">ALT+TAB</td>
<td class="org-left">Switch App</td>
<td class="org-left">Switch between open applications</td>
</tr>

<tr>
<td class="org-left">SHIFT+ALT+TAB</td>
<td class="org-left">Switch App Reverse</td>
<td class="org-left">Switch between open applications in reverse</td>
</tr>
</tbody>
</table>


<a id="org53c35d7"></a>

# Modifiers for Movement and Selection

<table border="2" cellspacing="0" cellpadding="6" rules="groups" frame="hsides">


<colgroup>
<col  class="org-left" />

<col  class="org-left" />

<col  class="org-left" />
</colgroup>
<tbody>
<tr>
<td class="org-left"><b>Key</b></td>
<td class="org-left"><b>Behavior</b></td>
<td class="org-left">*Description</td>
</tr>

<tr>
<td class="org-left">SHIFT</td>
<td class="org-left">Selection</td>
<td class="org-left">Expands object selection to span the range</td>
</tr>

<tr>
<td class="org-left">CTRL</td>
<td class="org-left">Macro (Big) Movement</td>
<td class="org-left">Enables Macro (bigger) movements</td>
</tr>
</tbody>
</table>

