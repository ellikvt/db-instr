# db-instr

Commissioning and Start-Up Instrumentation managing Database project
---------------------------------------------------
---------------------------------------------------

2025-10-28 update: File 'Информатор КИП v0.0.2_2014-11-17-date-bug-fxd' to be used instead of  
'Информатор КИП v0.0.2_2014-11-17' on the reason of modified code  
(pls see commit 9b9598cfce89b48c10c1fe10b398e75de6eddefa)  

**bug:** Error 424 on click with 'Учетная карта ТМЦ' button of the Main View  
**bug:** Macro command Error on click with 'Открыть форму' button of the Main View  
**bug:** Runtime Error 3085 on click with 'Отчет по использованию' button of the Main View  
**bug:** Error of Date, Time required on click with 'Отчет по сессиям' of the Main View  
**bug:** Error of Report Form positioning on the screen. After appear its close control might be not visible.  
Access restoring only after client reconnect to the VM  
**bug:** TreeCtl Access Error on click 'Добавить пути из дерева' button of the Main View.  
And this causes the VBA Run-time Error 91 'Object Variable of With Block Variable not set'.  
After that the Application stops proper responding on buttons clicks with the same VBA Run-time Error.  
**bug:** If TreeCtl accessed by clicking rigth after the Relocate Form appeared then it causes the VBA Run-Time Error 91.  
As a result the Form is not able to be closed without the code run break  
**bug:** Relocate Form Date is not initialized by today date but initialized with some wrong '##' data instead  



