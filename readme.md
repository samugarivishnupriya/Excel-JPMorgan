# Excel Certification by JPMorgan Chase & co.

## Overview
In this program, there are 5 different tasks have been given to which we need to submit the solution. Let's see each task in detail.

### TASK 1 : Excel Keyboard Shortcuts

The task is to learn some key Excel shortcuts, practice them in Excel, and then take the quiz which contains 6 MCQs.

Start by referring to the list of shortcuts by watching some brief videos on shortcuts using the links below. Then, try the shortcuts we learned from those videos. The best way to learn them is to use them, open any spreadsheet in Excel, and try out the various keyboard shortcuts. We will quickly find the ones we like and want to use going forward.

*Spreadsheet actions and movement:*
```markdown
Save: Ctrl-S
Print: Ctrl-P
Undo: Ctrl-Z
Redo: Ctrl-Y
Jump to Bottom of Data: Ctrl-Down Arrow
Jump to Top of Data: Ctrl-Up Arrow
Go to Previous Sheet: Ctrl-Page Up
Go to Next Sheet: Ctrl-Page Down
```
*Data selection:*
```markdown
 Select All Data in Column: Ctrl-Shift-Down Arrow
 Select Whole Column: Ctrl-Space
 Select All Data in Row: Shift-Space
 Select All Data in Region: Ctrl-A
```
*Data Editing:*
```markdown
 Filter: Ctrl-Shift-L
 Find: Ctrl-F
 Flash Fill: Ctrl-E
 Select Discontinuous Cells:  Shift-F8
 Spell Check: F7
 Automatically Add Sum Totals for Columns and/or Rows: Alt-=
 Remove Duplicates: Alt-A-M
 Sort A-Z: Alt-A-S-A
 Sort Z-A: Alt-A-S-D
 Delete Row: Ctrl-minus
 Copy Selected Object or Data: Ctrl-D
```
*Data Formatting:*
```markdown
 Format Cells: Ctrl-1
 Create Data Table: Ctrl-T
 Autoformat Data Tables: Alt-O-A
 Format As Currency: Ctrl-Shift-4
 Format As 2 Decimal Place Number: Ctrl-Shift-1
```
These are the links that we can refer to are :
1. [Excel Keyboard Shortcuts used by Pros](https://www.youtube.com/watch?v=4xanM8XD058)
2. [Excel shortcut keys you SHOULD know!](https://www.youtube.com/watch?v=Xe4U_-o_EWw)
3. [10 Excel Keyboard Shortcuts](https://www.youtube.com/watch?v=c2LHiAwxTt0)


### TASK 2 : Conditional Formatting

The task is to use Excel's conditional formatting tools to explore and visualize the characteristics of the data in the dataset.

Conditional formatting is a technique in Excel to make data exploration visual. We format the data (its colors, fonts, highlighting, etc.) based on the criteria we select. This allows us to quickly identify values that are either a data quality problem (missing values, formula errors, nonsensical values, etc.) or a value that drives insights and decisions (i.e., sales growth below a set target; inventory below a set amount, which would trigger re-ordering; etc.)

Excel's conditional formatting tools are easy to use but have a large and powerful array of options for criteria to base formatting on, including the ability to write our formulas. We can even use conditional formatting tools to format colorful and visually easy-to-read reports and simple dashboards of our data, creating visuals such as "heat maps" and even "Harvey ball" rankings to visually indicate data that meet specific criteria.

First, we need to familiarize ourselves with Excel's conditional formatting tools by watching the introductory videos using the links provided below.

Then, open the spreadsheet `Account Sales Data for Analysis` and familiarize ourselves with the data. What kind of data is there? What information do the columns contain? What kind of trends could we see with this kind of data?

Then, use the conditional formatting tools (either the menu-based tools or our conditional formatting formulas, whichever we prefer) to do the following explorations of the data:
```markdown
1. Highlight any cells with formula errors in purple with white text.
2. Highlight any cells with missing values in yellow.
3. Identify accounts that have not been cross-sold with Product 2 by highlighting the appropriate Product 2 cells in orange.
4. Identify accounts that have a 5-year sales CAGR of at least 100% by highlighting the appropriate CAGR cells in green and any
   accounts with a negative CAGR in red with white text.
5. Identify accounts in the top 10% of unit sales for 2021 by highlighting the appropriate 2021 unit sales cells in blue.
```
> For the solution check the `Account Sales Data solution for task 2`.

Please refer to the below references:
1. [Conditional Formatting in Excel Tutorial](https://www.youtube.com/watch?v=Jp29JYGq5Hw)
2. [Excel Conditional Formatting in Depth](https://www.youtube.com/watch?v=aomParo6vZ0)
3. [8 Expert Tricks for Conditional Formatting in Excel](https://www.youtube.com/watch?v=lIqifDg2xfE)

### TASK 3 : Visual Basic for Applications (VBA) Macros

Visual Basic for Applications (VBA) is a useful language to learn and skill to develop because it is built into the entire suite of Microsoft products, meaning VBA is also the programming language built into Word, PowerPoint, and other Microsoft applications. This means we can use VBA to integrate data and reporting across the entire Microsoft suite.

Like any programming language, VBA has its vocabulary, syntax, and commands to learn. With VBA, we can write sophisticated programs that exceed tasks we could do manually on the keyboard and mouse. While becoming skilled at writing complex VBA programs takes many hours of training, we can learn in just a few minutes to create simple macros to automate common, repetitive tasks in Excel.

A macro is simply a short list of commands written in VBA to automate a set of tasks we could otherwise do manually using the keyboard and mouse. Excel has two methods built-in for creating macros. The easiest way is to "record" the macro, which means telling Excel to"watch" our actions as we do a task using the keyboard and mouse and automatically create a list of commands in VBA that correspond to those actions. Then, we can tell Excel to automatically run that list of commands over and over as needed. That list or script of commands is our macro, and we can assign it to a button on the screen to run it at will. The other method for creating macros is to write the list of commands (the VBA code) ourselves without having Excel watch our actions and generate that list for us.

The task is to familiarize ourselves with recording and using simple macros in Excel, and then create two macros using the same spreadsheet `Account Sales Data for Analysis`.

In most versions of Excel, we have to enable the "developer" tab in the menu to work with macros and VBA. If our copy of Excel does not have this tab visible, we can right-click on the ribbon and add the tab to the menu.
```markdown
1. A macro to sort the entire spreadsheet by 5 YR CAGR in descending order to see which accounts have the highest overall 5-year sales growth
2. A macro to sort the entire spreadsheet by 2021 unit sales in descending order to see which accounts have the highest overall unit sales in 2021
```
When we are finished, we will have two buttons that let us very quickly and easily see two ways of analyzing account sales data to inform account planning and other operational decision-making and quickly switch between them.
> For a solution check the `Account Sales Data solution for Task 3`.

The below figure shows how to macro can be created.

![record macro](https://github.com/samugarivishnupriya/Excel-JPMorgan/assets/85831285/e1d8156d-4230-44b3-8d2b-3aa7187fffeb)

For recording the macro, follow the steps given in the below references.

1. [Beginners Guide to Excel Macros ](https://www.youtube.com/watch?v=wBDp9G2zWe8)
2. [Excel VBA - Record a Macro](https://www.youtube.com/watch?v=ltcpaHdXUrU)
3. [Excel VBA - Write a Simple Macro](https://www.youtube.com/watch?v=PoIVp9VWo4I)

### TASK 4: Data Visualization in Excel
To make data-driven business decisions, the decision-makers need an easy way to understand and draw conclusions about insights from our data and analysis. Insights can be financial, operational, or related to any other management need. One of the easiest ways to make such insights quickly understandable is by using charts or graphs of the data and our analysis of it.

Business decisions that need the same insights routinely benefit from an interactive or dynamic combination of charts called a dashboard. Dashboards allow us to see different views or "slices" of the data, see how different insights relate to each other, and gain a complete picture of what the data is saying. This is particularly helpful when making operational decisions like which products to market more or differently, which accounts need more sales activity to drive more sales, or which products need better inventory management to keep in stock. A dashboard is simply a collection of related charts on one page to make visualizing the data easier.

The task is to create a simple dashboard using the account sales dataset `Account Sales Data for Analysis`. Then, consider the dataset. What charts and graphs would be useful related to this data? Unit sales by year? Top 10 accounts by unit sales or CAGR? Effectiveness of different marketing programs by the number of sales driven? Sales by account type? There are a variety of different ways we could gain insight from this dataset. Pick the ones we find most compelling, and use those to create our dashboard.

Next, consider how we may need to transform the data in the dataset to simplify our analyses. Raw data is as we find it and often not in the ideal form for analysis. We may need to alter the spreadsheet structure or add calculations to support our analysis. Hint: disaggregating the raw data by building a new sheet that has a row per sales year per account, rather than a row per account that combines sales data for all five years, will make it much easier to use pivot tables for some types of analysis. Could we use one or more macros to make constructing that new sheet easier? we may want to filter the data into different views. We will want to add pivot tables to support some kinds of charts we could create. Feel free to change the dataset in any way that supports our analysis.

Make the data an Excel table (rather than a range). Remember the shortcut for that? It is Ctrl-T. Some of Excel's more useful capabilities work with data designated as a Table in Excel, including dynamic updating of charts and graphs and much of the pivot table functionality. It is a best practice to use Excel Tables when doing data analysis.
> For a solution check the `Account Sales Data solution for Task 4`

For creating the Pivot tables, see the below references:
1. [Create charts in Excel](https://www.excel-easy.com/data-analysis/charts.html)
2. [How to build a graph in Excel](https://blog.hubspot.com/marketing/how-to-build-excel-graph)
3. [CAGR formula in Excel](https://www.excel-easy.com/examples/cagr.html)
4. [Secrets to Building Excel Dashboards](https://www.youtube.com/watch?v=9p6tWCHbtPQ)
5. [How to Create Dashboards in Excel](https://www.youtube.com/watch?v=JcdORXZjbbg)

### TASK 5 : Data-Driven Storytelling
Telling a data-driven narrative allows us to directly link the issue we want to address, the insights useful to address the issue, and the action or decision we want the reader or listener to take. This linkage depends on the listener's emotional connection to the story.  The data alone does not convey an emotional connection; the narrative does that.

Telling a story with our data makes it easier to build trust, convey meaningful insight, and drive audience engagement.

Follow certain best practices when telling a data-driven story:

* Understand our audience. Who are we communicating with? What are their motivations and needs? What will they find compelling?
* Focus on a few major points. Keep our messaging clear and concise, so it’s memorable.
* Set the context for our story. Why do our insights matter? Why should the audience care?
* write a story. Stories include what is commonly called a story arc: a setup, a tension or issue, a resolution, and/or a
  call to action.
* Use visuals where possible. Tables of numbers are hard to make compelling.
* Support our credibility as the storyteller. Be honest about data quality issues, missing insights that are needed, and related
  risks to the action or decision we are seeking.

The task is to write a short PowerPoint presentation to communicate key insights and data from our analysis and visualization work in the prior task. From that work, we have insights into which accounts are and are not performing well, how sales are growing year over year, which account types are selling more units than others, and other kinds of findings that we could communicate. Use our analytical insights and even our Excel dashboard if we like as part of the written story we tell in PowerPoint.

First, do the background learning on data-driven storytelling using the links in Additional Resources below. Then, review our analysis and dashboard from Task 4. What data, ideas, insights, or examples would be compelling to decision-makers about account sales? Do other analyses of the dataset if needed to help us tell a compelling story.

Then, write our presentation in 3-5 slides, but remember: clear, concise, and compelling. Shorter is almost always better.

> For a solution check the `Data Driven Storytelling`

To do the PPT, refer to the below links.
1. [Best techniques for Data-Driven storytelling](https://www.gokantaloupe.com/blog/best-techniques-for-data-driven-storytelling)
2. [What is Data-Driven storytelling](https://www.revealbi.io/glossary/data-driven-storytelling)
3. [Data-Driven storytelling guide](https://unscrambl.com/blog/data-driven-storytelling-guide/)
4. [Why Data-Driven storytelling matters ](https://phrazor.ai/blog/the-art-of-data-driven-storytelling-what-is-it-and-why-does-it-matter)
5. [Storytelling in business](https://www.indeed.com/career-advice/career-development/storytelling-in-business)
6. [Business storytelling guide](https://www.lafabbricadellarealta.com/business-storytelling-the-definitive-guide/)
7. [Importance storytelling business](https://virtualspeech.com/blog/importance-storytelling-business)
