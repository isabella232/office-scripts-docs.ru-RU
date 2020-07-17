---
title: Поддержка старых сценариев Office, использующих асинхронные API
description: Знакомство с асинхронными API сценариев Office и использование шаблона нагрузки/синхронизации для старых сценариев.
ms.date: 07/08/2020
localization_priority: Normal
ms.openlocfilehash: e7ca5b276cff0e3a38bffc2af1541c0051cf5490
ms.sourcegitcommit: ebd1079c7e2695ac0e7e4c616f2439975e196875
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/17/2020
ms.locfileid: "45160462"
---
# <a name="support-older-office-scripts-that-use-the-async-apis"></a>Поддержка старых сценариев Office, использующих асинхронные API

В этой статье описывается поддержка и обновление скриптов, использующих асинхронные интерфейсы API модели предыдущих версий. Эти API имеют те же основные функциональные возможности, что и стандартные, синхронные API сценариев Office, но для управления синхронизацией данных между сценарием и книгой требуется ваш сценарий.

> [!IMPORTANT]
> Асинхронную модель можно использовать только со скриптами, созданными до реализации текущей [модели API](scripting-fundamentals.md?view=office-scripts). Скрипты окончательно блокируются до модели API, созданной им после создания. Это также означает, что если вы хотите преобразовать старый скрипт в новую модель, необходимо создать новый сценарий. Мы рекомендуем обновлять старые сценарии в новой модели при внесении изменений, так как использование текущей модели упрощается. [Преобразование асинхронных скриптов в текущий раздел модели](#converting-async-scripts-to-the-current-model) содержит советы по переходу на эту стадию.

## <a name="main-function"></a>Функция `main`

Скрипты, использующие асинхронные API, имеют различные `main` функции. Это `async` функция, которая является `Excel.RequestContext` первым параметром.

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your async Office Script
}
```

## <a name="context"></a>Context

Функция `main` принимает `Excel.RequestContext` параметра с именем `context`. Думайте о `context` как о мосте между вашим сценарием и книгой. Ваш сценарий обращается к книге с помощью `context` объекта и использует этот `context` для отправки данных туда и обратно.

Объект `context` необходим, потому что скрипт и Excel работают в разных процессах и местах. Сценарий должен будет внести изменения или запросить данные из рабочей книги в облаке. Объект `context` управляет этими транзакциями.

## <a name="sync-and-load"></a>Синхронизация и загрузка

Поскольку ваш сценарий и рабочая книга работают в разных местах, любая передача данных между ними занимает много времени. В асинхронном API команды ставятся в очередь до тех пор, пока не будет явно вызвана `sync` операция синхронизации скрипта и рабочей книги. Ваш скрипт может работать независимо, пока он не выполнит одно из следующих действий:

- Прочитайте данные из рабочей книги (с помощью операции `load` или метода возвращения [ClientResult](/javascript/api/office-scripts/excelscript/excelscript.clientresult?view=office-scripts-async)).
- Запишите данные в рабочую книгу (обычно потому, что сценарий завершен).

На следующем рисунке показан пример потока управления между сценарием и книгой:

![Диаграмма, показывающая операции чтения и записи, идущие в рабочую книгу из сценария.](../images/load-sync.png)

### <a name="sync"></a>Синхронизировать

Когда сценарий Async должен считывать данные из книги или записывать данные в нее, вызовите метод, `RequestContext.sync` как показано ниже:

```TypeScript
await context.sync();
```

> [!NOTE]
> `context.sync()` неявно вызывается, когда скрипт заканчивается.

После завершения операции `sync` книга обновляется, чтобы отразить все операции записи, указанные сценарием. Операция записи задает свойство для объекта Excel (например, `range.format.fill.color = "red"` ) или вызывает метод, который изменяет свойство (например, `range.format.autoFitColumns()` ). Операция `sync` также считывает любые значения из рабочей книги, запрошенные сценарием с помощью операции `load` или метода возвращения `ClientResult` (как описано в следующих разделах).

Синхронизация вашего сценария с книгой может занять некоторое время, в зависимости от вашей сети. Минимизируйте количество `sync` вызовов, чтобы ускорить выполнение сценария. В противном случае асинхронные API не будут быстрее стандартными, синхронными API.

### <a name="load"></a>Load

Асинхронный скрипт должен загружать данные из книги, прежде чем считывать их. Однако загрузка данных из всей книги значительно сокращает скорость сценария. `load`Метод позволяет скрипту указать, какие данные следует извлечь из книги.

Метод `load` доступен для каждого объекта Excel. Ваш скрипт должен загрузить свойства объекта, прежде чем он сможет их прочитать. Это не приведет к ошибке.

В следующих примерах объект `Range` используется для демонстрации трех способов использования метода `load` для загрузки данных.

|Intent |Пример команды | Эффект |
|:--|:--|:--|
|Загрузить одно свойство |`myRange.load("values");` | Загружает одно свойство, в данном случае двумерный массив значений в этом диапазоне. |
|Загрузить несколько свойств |`myRange.load("values, rowCount, columnCount");`| Загружает все свойства из списка, разделенного запятыми, в этом примере значения, количество строк и количество столбцов. |
|Загрузить все | `myRange.load();`|Загружает все свойства в диапазоне. Это не рекомендуемое решение, так как оно замедляет выполнение скрипта, получая ненужные данные. Используйте его только при тестировании скрипта или при необходимости для каждого свойства объекта. |

Ваш скрипт должен вызывать `context.sync()` перед чтением любых загруженных значений.

```TypeScript
/**
 * This script uses the async API to get the row count for a range.
 * It shows how to load a property in the async model.
 */
async function main(context: Excel.RequestContext) {
    let selectedSheet = context.workbook.worksheets.getActiveWorksheet();
    let range = selectedSheet.getRange("A1:B3");

    // Load the property.
    range.load("rowCount");

    // Synchronize with the workbook to get the property.
    await context.sync();

    // Read and log the property value (3).
    console.log(range.rowCount);
}
```

Вы также можете загрузить свойства всей коллекции. Каждый объект Collection в асинхронном API имеет `items` свойство, которое представляет собой массив, содержащий объекты в этой коллекции. Использование `items` в качестве начала иерархического вызова (`items\myProperty`) для `load` загружает указанные свойства для каждого из этих элементов. В следующем примере загружается свойство `resolved` для каждых `Comment` объектов в `CommentCollection` объекте рабочего листа.

```TypeScript
/**
 * This script uses the async API to get resolved property on every comment in the worksheet.
 * It shows how to load a property from every object in a collection.
 */
async function main(context: Excel.RequestContext){
    let selectedSheet = context.workbook.worksheets.getActiveWorksheet();
    let comments = selectedSheet.comments;

    // Load the `resolved` property from every comment in this collection.
    comments.load("items/resolved");

    // Synchronize with the workbook to get the properties.
    await context.sync();
}
```

### <a name="clientresult"></a>ClientResult

Методы в асинхронном API, возвращающие сведения из книги, имеют похожий шаблон для `load` / `sync` парадигмы. Например, `TableCollection.getCount` получает количество таблиц в коллекции. `getCount`Возвращает значение `ClientResult<number>` , означающее, что `value` возвращаемое свойство [`ClientResult`](/javascript/api/office-scripts/excelscript/excelscript.clientresult?view=office-scripts-async) является числом. Скрипт не может получить доступ к этому значению, пока не вызовет `context.sync()`. По аналогии с загрузкой свойства, `value` — это локальное пустое значение до вызова `sync`.

Следующий сценарий получает общее количество таблиц в рабочей книге и записывает его в консоль.

```TypeScript
/**
 * This script uses the async API to get the table count of the workbook.
 * It shows how ClientResult objects return workbook information.
 */
async function main(context: Excel.RequestContext) {
    let tableCount = context.workbook.tables.getCount();

    // This sync call implicitly loads tableCount.value.
    // Any other ClientResult values are loaded too.
    await context.sync();

    // Trying to log the value before calling sync would throw an error.
    console.log(tableCount.value);
}
```

## <a name="converting-async-scripts-to-the-current-model"></a>Преобразование асинхронных скриптов в текущую модель

В текущей модели API не используется `load` , `sync` или `RequestContext` . Благодаря этому скрипты значительно упрощают процесс записи и обслуживания. Лучший ресурс для преобразования старых сценариев — [переполнение стека](https://stackoverflow.com/questions/tagged/office-scripts). В этом случае вы можете обратиться к сообществу для получения справки по определенным сценариям. Следующие рекомендации должны помочь в структурировании общих действий, которые необходимо выполнить.

1. Создайте новый скрипт и скопируйте в него старый асинхронный код. Не включайте старую `main` подпись метода, используя `function main(workbook: ExcelScript.Workbook)` вместо нее текущую.

2. Удаление всех `load` вызовов и `sync` вызовов. Они больше не нужны.

3. Удалены все свойства. Теперь вы получаете доступ к этим объектам `get` и `set` методам, поэтому вам потребуется переключить ссылки этих свойств на вызовы методов. Например, вместо настройки цвета заливки ячейки с помощью доступа к свойству, как показано ниже: `mySheet.getRange("A2:C2").format.fill.color = "blue";` , вы можете использовать такие методы, как:`mySheet.getRange("A2:C2").getFormat().getFill().setColor("blue");`

4. Классы коллекций заменены на массивы. `add`Методы и для `get` этих классов коллекций были перемещены в объект, который владеет коллекцией, поэтому ваши ссылки должны быть соответствующим образом обновлены. Например, чтобы получить диаграмму с именем "myChart устанавливается подпись" на первом листе книги, используйте следующий код: `workbook.getWorksheets()[0].getChart("MyChart");` . Обратите внимание на то, что `[0]` нужно получить доступ к первому значению, `Worksheet[]` возвращенному методом `getWorksheets()` .

5. Некоторые методы были переименованы для ясности и добавлены для удобства. Дополнительные сведения см. в [справочнике по API сценариев Office](/javascript/api/office-scripts/overview?view=office-scripts) .

## <a name="office-scripts-async-api-reference-documentation"></a>Справочная документация по асинхронному API для сценариев Office

[!INCLUDE [Async reference documentation](../includes/async-reference-documentation-link.md)]
