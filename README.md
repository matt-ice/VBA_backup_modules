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
  -Currently has 2 functions: Trim and Restore
  
Trim:  
  -Applies TRIM() function to either selected cells (more than 2) or all used range (if only 1 cell is selected)  

Restore:  
  -Restores Screen updating, Calculation, Events and sets Interactive
