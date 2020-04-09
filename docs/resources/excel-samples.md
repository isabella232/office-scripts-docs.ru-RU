---
title: Примеры сценариев для сценариев Office в Excel в Интернете
description: Коллекция примеров кода для использования со сценариями Office в Excel в Интернете.
ms.date: 04/06/2020
localization_priority: Normal
ms.openlocfilehash: abf6b87b63ad027cca8ee5c947b687f54815409c
ms.sourcegitcommit: 0b2232c4c228b14d501edb8bb489fe0e84748b42
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/08/2020
ms.locfileid: "43191010"
---
# <a name="sample-scripts-for-office-scripts-in-excel-on-the-web-preview"></a>Примеры сценариев для сценариев Office в Excel в Интернете (Предварительная версия)

Ниже приведены примеры простых сценариев, которые можно использовать в собственных книгах. Чтобы использовать их в Excel в Интернете, выполните следующие действия:

1. Откройте вкладку **Автоматизировать**.
2. Нажмите клавишу **Редактор кода**.
3. Нажмите **новый скрипт** в области задач редактора кода.
4. Замените весь сценарий выбранным образцом.
5. В области задач редактора кода нажмите кнопку **запустить** .

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="scripting-basics"></a>Основные сведения о сценариях

В этих примерах демонстрируются основные конструктивные блоки для сценариев Office. Добавьте их в скрипты, чтобы расширить решение и устранить распространенные проблемы.

### <a name="read-and-log-one-cell"></a>Чтение и запись в журнал одной ячейки

В этом примере считывается значение **a1** и выводится на консоль.

``` TypeScript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Get the value of cell A1.
  let range = selectedSheet.getRange("A1");
  range.load("values");
  await context.sync();

  // Print the value of A1.
  console.log(range.values);
}
```

### <a name="work-with-dates"></a>Работать с датами

В примерах, приведенных в этом разделе, показано, как использовать объект JavaScript [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) .

В следующем примере возвращается текущая дата и время, а затем эти значения записываются в две ячейки активного листа.

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the cells at A1 and B1.
  let dateRange = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
  let timeRange = context.workbook.worksheets.getActiveWorksheet().getRange("B1");

  // Get the current date and time with the JavaScript Date object.
  let date = new Date(Date.now());

  // Add the date string to A1.
  dateRange.values = [[date.toLocaleDateString()]];
  
  // Add the time string to B1.
  timeRange.values = [[date.toLocaleTimeString()]];
}
```

В следующем примере считывается дата, которая хранится в Excel, и преобразуется в объект даты JavaScript. В качестве входных данных для даты JavaScript в качестве входных данных используется [числовой серийный номер даты](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) .

```TypeScript
async function main(context: Excel.RequestContext) {
  // Read a date at cell A1 from Excel.
  let dateRange = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
  dateRange.load("values");
  await context.sync();

  // Convert the Excel date to a JavaScript Date object.
  let excelDateValue = dateRange.values[0][0];
  let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
  console.log(javaScriptDate);
}
```

## <a name="display-data"></a>Отображение данных

В этих примерах показано, как работать с данными листа и предоставлять пользователям лучшее представление или организацию.

### <a name="apply-conditional-formatting"></a>Применение условного форматирования

В этом примере применяется условное форматирование для диапазона, используемого в текущий момент на листе. Условное форматирование — Зеленая заливка для первых 10% значений.

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Get the used range in the worksheet.
  let range = selectedSheet.getUsedRange();

  // Set the fill color to green for the top 10% of values in the range.
  let conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.topBottom);
  conditionalFormat.topBottom.format.fill.color = "green";
  conditionalFormat.topBottom.rule = {
    rank: 10, // The percentage threshold.
    type: Excel.ConditionalTopBottomCriterionType.topPercent // The type of the top/bottom condition.
  };
}
```

### <a name="create-a-sorted-table"></a>Создание отсортированной таблицы

В этом примере создается таблица на основе используемого диапазона текущего листа, а затем она сортируется по первому столбцу.

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Create a table with the used cells.
  let usedRange = selectedSheet.getUsedRange();
  let newTable = selectedSheet.tables.add(usedRange, true);

  // Sort the table using the first column.
  newTable.sort.apply([{ key: 0, ascending: true }]);
}
```

## <a name="collaboration"></a>Совместная работа

В этих примерах показано, как работать с функциями Excel, относящимися к совместной работе, например комментариями.

### <a name="delete-resolved-comments"></a>Удаление разрешенных комментариев

В этом примере удаляются все разрешенные комментарии из текущего листа.

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Get the comments on this worksheet.
  let comments = selectedSheet.comments;
  comments.load("items/resolved");
  await context.sync();

  // Delete the resolved comments.
  comments.items.forEach((comment) => {
      if (comment.resolved) {
          comment.delete();
      }
  });
}
```

## <a name="scenario-samples"></a>Примеры сценариев

Примеры, иллюстрирующие большие, реальные решения, можно найти на странице [примеры сценариев для сценариев Office](scenarios/sample-scenario-overview.md).

## <a name="suggest-new-samples"></a>Предлагаемые новые примеры

Мы будем рады получать новые примеры. Если существует распространенный сценарий, который поможет другим разработчикам скриптов, Расскажите нам в разделе отзывов, приведенном ниже.
