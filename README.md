# db-instr

Commissioning and Start-Up Instrumentation managing Database project
---------------------------------------------------
---------------------------------------------------

**2025-10-28 update:** File 'Информатор КИП v0.0.2_2014-11-17-date-bug-fxd' to be used instead of  
'Информатор КИП v0.0.2_2014-11-17' on the reason of modified code  
(pls see commit 9b9598cfce89b48c10c1fe10b398e75de6eddefa)  
**2025-11-07 update:** Binary File 0.0.2-bugs-#1,9-fixed to be used (pls see commit af30cba)  

`**bug 1:** fixed` Error 424 on click with 'Учетная карта ТМЦ' button of the Main View  
`**bug 2:** fixed` Macro command Error on click with 'Открыть форму' button of the Main View  
( fixed in Binary File 0.0.2-bugs-#1,9-fixed by removing button related not correct macro command )  
**bug 3:** Runtime Error 3085 on click with 'Отчет по использованию' button of the Main View  
**bug 4:** Error of Date, Time required on click with 'Отчет по сессиям' of the Main View  
**bug 5:** Error of Report Form positioning on the screen. After appear its close control might be not visible.  
Access restoring only after client reconnect to the VM  
**bug 6:** TreeCtl Access Error on click 'Добавить пути из дерева' button of the Main View.  
And this causes the VBA Run-time Error 91 'Object Variable of With Block Variable not set'.  
After that the Application stops proper responding on buttons clicks with the same VBA Run-time Error.  
**bug 7:** If TreeCtl accessed by clicking rigth after the Relocate Form appeared then it causes the VBA Run-Time Error 91.  
As a result the Form is not able to be closed without the code run break  
**bug 8:** ViewСклад Form not capable to refresh automatically on relocation the remaining last selected node item.  
         This leads to Form stuck after relocation finished.  
`**bug 9:** fixed` Private Sub Form Load of Relocation Form contains Date field that is initialized with no proper '###' data on Form opening  
**bug 10:** Forcing ViewСклад Form close affects the other users. To chekc if that can be solved by dividing the Front-end App and Database



