I want to build a program using VsCode:



First, I want it able to run just through VsCode, that's sufficient enough, but next maybe turn it into openable program (application), like a simple one. Because I might encounter this kind of situation I need to use this many times, so it should be reusable.



So the flow of the system:



1\. First it read the excel file, then after that process it to turn it into something like this:



I sort it by column order which need to go first and explain each:



1. Organisation Name - I want it to grouped by this.
2. Parent company - 1 organisation should only have 1 type of parent company right?
3. Organisation Type - 1 organisation should only be 1 type of organisation right?
4. Organisation Size - 1 organisation should only be 1 type of organisation size?
5. Stakeholder Category - 1 organisation should only be 1 type of stakeholder category right?
6. PDCS Sector - 1 organisation should only be 1 type of PDCS sector right?
7. 1 organisation should only be located in 1 specific district right?
8. Role level - There could be multiple role level, but it should all be displayed in this column
9. Department - There could be multiple department, but it should all be displayed in this column
10. Age band - There could be multiple age band, but it should all be displayed in this column
11. Part of group - 1 organisation should only have 1 specific status for Part of group right?
12. Job Title - There could be multiple Job title, but it should all be displayed in this column
13. Section - this is also called dimension, this also could be the important pivot on what I want to see. Obviously there will be multiple sections of aspects you want to obtain from a company, but then it should all be listed in different rows for a specific section, but then consider grouped by this, but refer to the next order
14. Question ID - Referring above, for a specific section, there could be multiple Questions for that specific section, but what we want to see and measure is for each question, how a certain company answers, so this column should be Question ID, with list of rows of that Question ID
15. Question - So this column is the Question iself which tied to the question, so same case, it should displayed correctly on what Question ID it is, and what Question itself is
16. Answer - Same case for Question, this one is reflected based on whats the question, but then what if there's different answer from each participant? how do we cover that? should it be list of answers for 1 specific question, but then how I can display it in a way user understand?
17. Answer value - this one, tied to the answer itself, an answer have specific label for its answer value, but should be displayed here
18. Answer score - This one, is tied to Answer and Answer value itself, each answer has its own score, so it should be displayed and the value should be respective to the answer



Now, I hope you get what I mean, this is basically the gist idea of what I want to do.



Maybe for the first version, I want code that able to cater to do this in vscode environment itself. But, should also have the option for me to exclude certain column I don't want, or maybe also include certain column that I want, like maybe by commenting only a certain line, I'm not very sure either, I'm not an either, very beginner, please take that into account, so maybe some code structure that's easier for me to do and manage this way. But then, after running the code, it should produce an output excel file for this first version and display the result that I want correctly.



And for the second version, maybe an app, you know, that has GUI and things like that, but when let's say user upload an excel file, it will display the full table below in the app, and there will be space on top for button or options, or whatever you call it, to perform a specific filter that you want, like you want to group by this column, or this certain column, and maybe you want to exclude or include this column, something like that.



What's your advise on this? Is there improvisation or better vision you see with this? suggest and advise



Please, let's develop this, step by step, dont throw all steps in one window answer (because it would be too long and laggy), but dont cover too little step, lets say only 1 step at a time, ask me to say "Next" before proceed. Assume im a beginner.

