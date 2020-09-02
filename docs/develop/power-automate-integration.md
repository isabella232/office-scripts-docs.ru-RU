---
title: Запуск сценариев Office с помощью автоматизации управления питанием
description: Как получить скрипты Office для Excel в Интернете, работая с рабочими процессами Power Автоматизация.
ms.date: 07/24/2020
localization_priority: Normal
ms.openlocfilehash: 87bd4e15ef7680a7456077494e3fda8208d6b9d8
ms.sourcegitcommit: e9a8ef5f56177ea9a3d2fc5ac636368e5bdae1f4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/01/2020
ms.locfileid: "47321574"
---
# <a name="run-office-scripts-with-power-automate"></a>Запуск сценариев Office с помощью автоматизации управления питанием

[Power автоматизиру](https://flow.microsoft.com) позволяет добавлять скрипты Office в более крупный автоматизированный рабочий процесс. Вы можете использовать автоматизацию управления питанием, например добавить содержимое электронной почты в таблицу листа или создать действия в средствах управления проектами на основе комментариев к книгам.

## <a name="getting-started"></a>Начало работы

Если вы впервые используете автоматизированное управление питанием, рекомендуем [ознакомиться со статьей "начать автоматизацию](/power-automate/getting-started)". Здесь вы можете узнать больше о всех возможных возможностях автоматизации. В этом разделе приведены сведения о том, как сценарии Office работают с автоматизацией управления питанием и как они могут помочь повысить удобство работы с Excel.

Чтобы приступить к объединению сценариев Power автоматизированного и Office, следуйте рекомендациям, описанным в разделе [starting Scripts with Power Автоматизация](../tutorials/excel-power-automate-manual.md). С его помощью вы узнаете, как создать последовательность, вызывающую простой сценарий. После выполнения этого руководства и [передачи данных сценариям в руководстве автоматизированного управления энергопотреблением](../tutorials/excel-power-automate-trigger.md) вернитесь сюда для получения подробных сведений о подключении сценариев Office к автоматизации потоков Power автоматизированного.

## <a name="excel-online-business-connector"></a>Соединитель Excel Online (Business)

[Соединители](/connectors/connectors) — это мосты между автоматизированной автоматизацией и приложениями. [Соединитель Excel Online (Business)](/connectors/excelonlinebusiness) предоставляет потокам доступ к книгам Excel. Действие "Запуск скрипта" позволяет вызывать любой сценарий Office, доступный через выбранную книгу. Вы также можете предоставить входные параметры скриптов, чтобы данные могли быть предоставлены с помощью этого процесса, или чтобы скрипт возвращал сведения для последующих шагов в этом блоке.

> [!IMPORTANT]
> Действие "Запуск скрипта" дает пользователям, использующим Microsoft Connector, значительный доступ к книге и ее данным. Кроме того, существуют риски, связанные с безопасностью, с помощью скриптов, которые выполняют внешние вызовы API, как описано во [внешних вызовах от автоматизации Powering](external-calls.md). Если администратор имеет дело с очень конфиденциальными данными, он может либо отключить Microsoft Excel Online Connector, либо ограничить доступ к сценариям Office с помощью [сценариев Office](/microsoft-365/admin/manage/manage-office-scripts-settings).

## <a name="data-transfer-in-flows-for-scripts"></a>Передача данных в потоках для сценариев

Power автоматизиру позволяет передавать фрагменты данных между действиями в рамках своего процесса. Сценарии можно настроить так, чтобы они принимали нужные вам типы сведений и возвращать что угодно из вашей книги, которую вы хотите использовать в вашем движении. Входные данные для вашего скрипта задаются путем добавления параметров в `main` функцию (в дополнение к `workbook: ExcelScript.Workbook` ). Выходные данные сценария объявляются путем добавления типа возвращаемого значения в `main` .

> [!NOTE]
> При создании блока "выполнить скрипт" в потоке заполняются допустимые параметры и возвращаемые типы. Если вы изменяете параметры или типы возвращаемых данных в вашем сценарии, вам потребуется повторить блок потока "Run script". Это гарантирует, что данные анализируются правильно.

В следующих разделах рассматриваются входные и выходные данные для сценариев, используемых в автоматизации Powering. Если вы хотите получить практический подход к освоению этой статьи, ознакомьтесь со статьей " [Передача данных в скрипты в руководстве по потоку автоматизированного управления питанием](../tutorials/excel-power-automate-trigger.md) " или изучите пример сценария [автоматизированной задачи "напоминания](../resources/scenarios/task-reminders.md) ".

### <a name="main-parameters-passing-data-to-a-script"></a>`main` Параметры: передача данных в скрипт

Все входные данные сценария указываются как дополнительные параметры `main` функции. Например, если вы хотите, чтобы сценарий принимал объект `string` , представляющий имя в качестве входных данных, вы можете изменить `main` подпись на `function main(workbook: ExcelScript.Workbook, name: string)` .

Когда вы настраиваете потоки в Power Автоматизация, вы можете указать входные данные скрипта в виде статических значений, [выражений](/power-automate/use-expressions-in-conditions)или динамического содержимого. Подробные сведения о соединителе отдельных служб можно найти в [документации Power автоматизиру Connector](/connectors/).

При добавлении входных параметров в функцию сценария `main` учитывайте следующие ограничения и ограничения.

1. Первый параметр должен иметь тип `ExcelScript.Workbook` . Имя параметра не имеет значения.

2. Каждый параметр должен иметь тип (например, `string` или `number` ).

3. Основные типы,,,,,, `string` `number` `boolean` `any` `unknown` `object` и `undefined` поддерживаются.

4. Массивы приведенных выше базовых типов поддерживаются.

5. Вложенные массивы поддерживаются в качестве параметров (но не как типы возвращаемого значения).

6. Типы Union разрешены, если они являются объединением литералов, принадлежащих одному типу (например, `"Left" | "Right"` ). Также поддерживаются объединения поддерживаемого типа с неопределенной версией (например, `string | undefined` ).

7. Типы объектов разрешены, если они содержат свойства типа `string` , `number` , `boolean` , поддерживаемых массивов или других поддерживаемых объектов. В следующем примере показаны вложенные объекты, которые поддерживаются как типы параметров:

    ```TypeScript
    // Office Scripts can return an Employee object because Position only contains strings and numbers.
    interface Employee {
        name: string;
        job: Position;
    }

    interface Position {
        id: number;
        title: string;
    }
    ```

8. Объекты должны иметь определение интерфейса или класса, определенное в сценарии. Объект также может быть определен анонимно, как показано в следующем примере:

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

9. Необязательные параметры разрешены и могут быть отмечены с помощью необязательного модификатора `?` (например, `function main(workbook: ExcelScript.Workbook, Name?: string)` ).

10. Допустимые значения параметров по умолчанию (например `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')` ,.

### <a name="returning-data-from-a-script"></a>Возвращение данных из скрипта

Скрипты могут возвращать данные из книги для использования в качестве динамического контента в автоматизированном блоке управления питанием. Как и в случае с входными параметрами, Автоматизация управления питанием применяет некоторые ограничения к типу возвращаемого значения.

1. Поддерживаются основные типы,,,, `string` `number` `boolean` `void` и `undefined` .

2. Типы объединения, используемые в качестве возвращаемых типов, действуют с теми же ограничениями, что и при использовании в качестве параметров сценария.

3. Типы массивов разрешены, если они имеют тип `string` , `number` или `boolean` . Они также разрешены, если тип является поддерживаемым объединением или поддерживаемым типом литерала.

4. Типы объектов, используемые в качестве возвращаемых типов, действуют с теми же ограничениями, что и при использовании в качестве параметров сценария.

5. Неявная типизация поддерживается, несмотря на то, что они должны следовать тем же правилам, что и определенный тип.

## <a name="avoid-using-relative-references"></a>Избегайте использования относительных ссылок

Power автоматизирует выполнение вашего сценария в выбранной книге Excel от вашего имени. В этом случае книга может быть закрыта. Любой API, зависящий от текущего состояния пользователя (например `Workbook.getActiveWorksheet` ,), не будет работать при использовании автоматизации Powering. При проектировании скриптов обязательно используйте абсолютные ссылки на листы и диапазоны.

Приведенные ниже методы вызовут ошибку и завершатся ошибкой при вызове из скрипта в блоке автоматизации Power.

| Класс | Метод |
|--|--|
| [Chart](/javascript/api/office-scripts/excelscript/excelscript.chart) | `activate` |
| [Range](/javascript/api/office-scripts/excelscript/excelscript.range) | `select` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveCell` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveChart` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveSlicer` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRange` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRanges` |
| [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `activate` |

## <a name="example"></a>Пример

На следующем снимке экрана показан процесс автоматизации Power, который срабатывает при назначении вопроса [GitHub](https://github.com/) . Поток выполняет сценарий, который добавляет ошибку в таблицу в книге Excel. Если в этой таблице имеется пять или более проблем, посылается напоминание по электронной почте.

![Пример процесса, показанный в редакторе автоматизации управления питанием.](../images/power-automate-parameter-return-sample.png)

`main`Функция скрипта ЗАДАЕТ идентификатор вопроса и заголовок вопроса в качестве входных параметров, а скрипт возвращает количество строк в таблице "ошибка".

```TypeScript
function main(
  workbook: ExcelScript.Workbook,
  issueId: string,
  issueTitle: string): number {
  // Get the "GitHub" worksheet.
  let worksheet = workbook.getWorksheet("GitHub");

  // Get the first table in this worksheet, which contains the table of GitHub issues.
  let issueTable = worksheet.getTables()[0];

  // Add the issue ID and issue title as a row.
  issueTable.addRow(-1, [issueId, issueTitle]);

  // Return the number of rows in the table, which represents how many issues are assigned to this user.
  return issueTable.getRangeBetweenHeaderAndTotal().getRowCount();
}
```

## <a name="see-also"></a>См. также

- [Запуск сценариев Office в Excel в Интернете с помощью Power автоматизиру](../tutorials/excel-power-automate-manual.md)
- [Передача данных сценариям в автоматически запускаемых рабочих процессах Power Automate](../tutorials/excel-power-automate-trigger.md)
- [Основные сведения о сценариях Office в Excel в Интернете](scripting-fundamentals.md)
- [Начало работы с Power Automate](/power-automate/getting-started)
- [Справочная документация по Microsoft Online Connector (бизнес)](/connectors/excelonlinebusiness/)
