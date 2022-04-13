# ExcelFunctionRegex

## Описание
***ExcelFunctionRegex*** - расширение для Microsoft Excel написанное с использование библиотеки *Excel-DNA*, добавляющее в Excel возможность проверки, поиска и замены по регулярному выражению.

## Установка
К файлам расширения относятся файлы XLL, расположенные в каталоге [ExcelFunctionRegex/bin/Release/](ExcelFunctionRegex/bin/Release/). Понятно, что 64, в наименовании файла, означает 64-рёх битный Excel, а вот *packed* означает, что в этот файл упакованы какие-либо внешние файлы. В данном случае никакие дополнительные файлы при создании расшириения использованы не были, поэтому нет разницы который из фйлов использовать packed или нет, имеет значение только разрядность. 
Устанавливать их надо следующим образом:
- В параметрах Excel надо выбрать вкладку "*Надстройки*"; 
- В низу окна параметров Excel, при выбранном значении поля "*Надстройки Excel*", нажать кнопку "*Перейти...*". Появится небольшое окно со списком всех возможных надстроек в Excel.;
- Надо выбрать любую из надстроек, и нажать кнопку "*Обзор...*", при этом совсем не обязательно эту надстройку подключать;
- Откроется диалоговое окно, в котором можно увидеть путь, где должны располагаться надстройки. Этот путь нужно запомнить, а лучше скопировать его в буфер;
- Надо закрыть все диалоговые окна Excel, да и сам Excel тоже следует закрыть;
- Надо открыть любую папку и по запомненному пути выйти в директорий, в котором хранятся расширения Excel, хотя, когда Вы этот директорий откроете, то он, скорее всего, будет ещё пустым;
- В этот директорий надо скопировать один из фалов расширение из каталога [ExcelFunctionRegex/bin/Release/](ExcelFunctionRegex/bin/Release/), в соответствии с разрядностью установленного офиса;
- После того как расширение окажется в папке, надо снова открыть "*Параметры Excel*" и во вкладке "*Надстройки*" кнопкой "*Перейти...*" открыть окно со списком надстроек. На этот раз в списке надстроек уже можно будет увидеть добавленное расширение *ExcelFunctionRegex Add-In*, которое надо активировать, поставив напротив него галочку.

## Использование

В списке категорий функций найти и выбрать раздел *ExcelFunctionRegex Add-In*, после этого появиться возможность выбрать одну из функций:
- ДелТекстПоРегВыр
- ПоискПоРегВыр
- ПроверкаПоРегВыр
- СцепТекстПоРазд
