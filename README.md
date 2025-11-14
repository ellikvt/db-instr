# db-instr

Commissioning and Start-Up Instrumentation managing Database project
---------------------------------------------------
---------------------------------------------------

**Some previous histrorizated changes related with app develop:**  
- Информатор КИП v0.0.2
   - переработан полностью код формы ввода ТМЦ
       - новая форма (и все ее объекты) больше жестко не связаны с источниками данных, а являются 
	     самостоятельными объектами, что позволило более гибко работать с данными,
	     получаемыми в форме от пользователя
       - Для всех элементов управления источники данных инициализируются
	     явным образом в коде формы, что более предпочтительно ( в основном с точки зрения будующего портирования
	     в другую СУБД). То есть максимальной реализации принципа: управление данными и сами данные отделены
   - пароль на код VBA снят (для удобства дальнейшей модернизации)
   - устранены баги в коде класса "clsRelocateClass": 
       - переменная serNumPribor в методах defineInsideQuery и defineOutsideQuery инициализировалась строкой "Is NULL" перед использованием ее в составлении выражения запроса. Это давало ошибку, т.к.
	     тип переменной строковый и правильно инициализировать ее пустой строкой.
   - в ходе переработки формы ввода найден баг:
       - баг не позволял ввод ТМЦ с кол-вом > 1. Это происходило из-за неверного составления запроса на вставку в таблицу Тип_Наличие (вместо поля с названием Тип_Код, использовалось название Тип).
	     Как бы то ни было - процедура вставки переработана, а старая функция множественной вставки удалена.

- Информатор КИП v0.0.1
   - каждое ТМЦ имеет набор обязательных аттрибутов, которые указываются при его перемещении
       - дата перемещения
       - кем перемещено
       - причина перемещения
       - куда перемещено
       - При указании этих данных
 делается проверка вводимой информации на соответствие определенным условиям (типам, маскам ввода)
   - При процессе перемещения реализован принцип полного контроля за его ходом. В частности
       - с возможностью исправить вводимые данные перед их записью в БД.
   - Каждая порция данных вводится и проверяется отдельно:
       - с помощью дочерней формы "Ввод данных об операции перемещения"
       - вызываемой нажатием кнопки "Начать перемещение..."
       - После корректного ввода всех данных на этой форме происходит 
         переход к вводу последней порции данных "Куда перемещено".
   - При перемещении ТМЦ проверяется корректность ввода количества перемещаемых ТМЦ:
       - при выдаче можно переместить не более кол-ва находящихся до перемещения на складе
       - остаток на складе автоматически декрементируется после перемещения на количество перемещенных ТМЦ
   - добавлен путь вывода отчетов: папка "Мои документы"
   - добавлен отчет по ТМЦ АСУ ТП (выводится на печать, а также в формат PDF)
   - улучшена работа с  ТМЦ, являющихся расходными (бобышки, штуцеры, барьеры, реле и т.д., не имеющие заводского или инвентарного номера), вводимые в БД во множественном числе:
       - При вводе в БД достаточно  просто указать количество таких ТМЦ. 
       - При вводе  таких ТМЦ в БД создает уникальный код для каждой единицы товара. 
       - Соответственно - перемещения ТМЦ происходят по одной штуке за раз, с указанием цели перемещения. 
       - Таким образом  можно однозначно отследить каждое ТМЦ 


**2025-10-28 update:** File 'Информатор КИП v0.0.2_2014-11-17-date-bug-fxd' to be used instead of  
'Информатор КИП v0.0.2_2014-11-17' on the reason of modified code  
(pls see commit 9b9598cfce89b48c10c1fe10b398e75de6eddefa)  
**2025-11-07 update:** Binary File 0.0.2-bugs-#1,9-fixed to be used (pls see commit af30cba)  
**2025-11-12 update:** Binary File 0.0.3 to be used (pls see commit about bug 3 fix)  
**2025-11-13 update:** Binary File 0.0.4 to be used (on the bug 4 fix)  
**2025-11-14 update:** Binary File 0.0.5 to be used (on the bug 6 fix)  
**2025-11-14 update:** Binary File 0.0.6 to be used (on the bug 11 fix)  
 
`**bug 1:** fixed` Error 424 on click with 'Учетная карта ТМЦ' button of the Main View  
`**bug 2:** fixed` Macro command Error on click with 'Открыть форму' button of the Main View  
( fixed in Binary File 0.0.2-bugs-#1,9-fixed by removing button related not correct macro command )  
`**bug 3:** fixed` Runtime Error 3085 on click with 'Отчет по использованию' button of the Main View  
couldn't resolve the statement >(Date()-365) And <=Date() and it was deleted. Now the use report assume just the whole time  
`**bug 4:** fixed` Error of Date, Time required on click with 'Отчет по сессиям' of the Main View. - Date/Time textboxes removed due to not needed.  
And the Report Form to be positioned at tne center of the screen. Now the Form is unaccessible to close - no options found for DoCmd.OpenReport  
that are related with Form positioning.  Default Window Mode is left.  
The caption `Total qty of sessions registered` added also on the Report Form at the bottom.  
The report exporting PDF forman replaced by RTF to get more Windows portability  
Same is done regarding the ICSS report  
`**bug 5:** fixed` Error of Report Form positioning on the screen. After appear its close control might be not visible.  
Access restoring only after client reconnect to the VM  - no more this behaviour detected. It's exposed to be related with   
the mistakes of reporting in the app that were fixed with bugs #1,2,3,4 handling  
`**bug 6:** fixed` TreeCtl Access Error on click 'Добавить пути из дерева' button of the Main View.  
And this causes the VBA Run-time Error 91 'Object Variable of With Block Variable not set'.  
After that the Application stops proper responding on buttons clicks with the same VBA Run-time Error.  
Bug is related with rsSlave recordset which doesn't exist on the nodes exporting end due to it's not initialized.
The detail reason is `Do Until rsMaster.EOF Then .. Loop` loop which never initializing rsSlave recordset on rsMaster proceeding finish.  
This leads to the error when trying to close non existing rsSlave on the export finish.  
This option is not good to use by app users. The option to be switched off because such a DB structure proceed might to be waste of time when  
doing with VBA. This is DB Admin main responsibility to make the proper structure operations.  
`**bug 7:** fixed` If TreeCtl accessed by clicking rigth after the Relocate Form appeared then it causes the VBA Run-Time Error 91.  
As a result the Form is not able to be closed without the code run break  
It's exposed to be fixed already on the bugs # 1,2,3,4 clearing. Now the response of the TreeCtl is blocked by app on the relocation proceeding that is right response in fact.  
**bug 8:** ViewСклад Form not capable to refresh automatically on relocation the remaining last selected node item.  
         This leads to Form stuck after relocation finished.  
`**bug 9:** fixed` Private Sub Form Load of Relocation Form contains Date field that is initialized with no proper '###' data on Form opening  
**bug 10:** Forcing ViewСклад Form close affects the other users. To chekc if that can be solved by dividing the Front-end App and Database  
`**bug 11:**` On Debug Button click with Auth Form the app get the `user8` credentials.  
For the full debug it's not enough and credentials re-assigned with `admin8` that gives more expanded right to make operations in the app. 
This is basically not the bug as it appears but need to be corrected for full debug abilities.  
**bug 12:** If TreeCtl accessed by clicking rigth after the Relocate Form data given to the app by user and no choice is made regarding the destination ot relocation ( no node clicked in TreeCtl ) then the app generating not proper reply that need to be corrected.  
The button `Сохранить изменения` shall not be activated until proper destination node is selected by user.   
**bug 13:** In the `` Form the field `` is not properly indicating the fact of relocation instead of field ``  


