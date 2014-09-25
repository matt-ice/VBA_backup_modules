VBA_backup_modules
==================

Backup modules/procedures for VBA

Tab cycle:  
  -FlipTab for going forward one tab       (Ctrl+Q)  
  -BackFlipTab for going backward one tab  (Ctrl+Shift+Q)  
  -Both functions cycle through the sheets (don't stop once they reach the last/first sheet) 

Useful Launcher (Ctrl+Shift+U)  
  -Required Microsoft Visual Basic for Application Extensibility library  
  -Requires Trust access to the VBA project object model checked as it creates a form and populates it with code  
  -Currently has 4 functions: Trim, Restore, Make Hyperlink and Break Hyperlink
  
Trim:  
  -Applies TRIM() function to either selected cells (more than 2) or all used range (if only 1 cell is selected)  

Restore:  
  -Restores Screen updating, Calculation, Events and sets Interactive  
  
Make Hyperlink:
  -Does what it says on the tin. Takes the selection, makes it into a hyperlink. Skips cells that don't have a valid link
  
Break Hyperlink:
  -The opposite of the above. Takes selection and deletes hyperlinks that were there. Leaves text behind
