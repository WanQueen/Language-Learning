# Language Learning

This program is made for language learning, which based on winform and excel.  
In order to use this program, you need to make the correct excel file, which can be correctly read by the program.  

Here is the example.  
| 한국어 (Learning Word) | English (Translation / Meaning) | Hint |
| :----: | :----: | :----: |
| 영국 | Britain | An Island country at west-Europe |
| 모자 | hat | something worn on head |

You need to input the right Learning Word to get to the next word.


## Log
#### version 0.1:
Finished most of the basic function: 
 - reading excel files
 - the simple typeset of Controls
 - judging whether correct answer
 - random word question

 #### version 0.1.1:
 - updated the rand word-choosing way, which is better than the last version
 - add all-word-finished page
 - bug found (1): when the title-row only contains number, the program will shut down

#### version 0.1.2:
 - fixed the bug (1)

#### version 0.1.3:
 - add icon into package
 - realigned mode-choosing page's buttons
 - now when exception happened at loading file, the MessageBox will be popped and program will be terminated
