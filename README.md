# moodlexml2xls
Convert Moodle XML export to Excel to select, sort and review questions.

When editing a question bank moodle provides only a quite limited UI. 
Tags are truncated in the display, selection is only by tag, and display columns cannot be customized.

This script uses the XML export because that has the full data set, allowing various export options by modifying the script. 

The current version creates a table with name, type, questiontext, grade and a column per tag.
