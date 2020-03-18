---
title: Основные сведения о сценариях для сценариев Office в Excel в Интернете
description: Сведения об объектной модели и другие основные сведения, которые необходимо изучить перед написанием сценариев Office.
ms.date: 01/27/2020
localization_priority: Priority
ms.openlocfilehash: 5a709c16e23c00ffc7ee7949a3cb11459dc2d530
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700351"
---
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web-preview"></a>Основные сведения о сценариях для сценариев Office в Excel в Интернете (Предварительная версия)

В этой статье представлены технические аспекты сценариев Office. Вы узнаете, как объекты Excel работают вместе и как редактор кода синхронизируется с книгой.

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="object-model"></a>Объектная модель

Чтобы ознакомиться с API Excel, необходимо знать, как компоненты книги связаны друг с другом.

- **Книга** содержит один или несколько **листов**.
- **Лист** предоставляет доступ к ячейкам с помощью объектов **Range** .
- **Диапазон** представляет группу смежных ячеек.
- **Диапазоны** используются для создания и размещения **таблиц**, **диаграмм**, **фигур**и других объектов визуализации данных или организации.
- **Лист** содержит коллекции объектов данных, присутствующих на отдельном листе.
- **Книги** содержат коллекции некоторых объектов данных (например, **таблиц**) для всей **книги**.

### <a name="ranges"></a>Ranges

Диапазон — это группа смежных ячеек в книге. В сценариях обычно используется нотация в стиле a1 (например, **B3** для одной ячейки в строке **B** и столбца **3** или **C2: F4** для ячеек из строк **C** и **F** и столбцов **2** – **4**) для определения диапазонов.

Диапазоны имеют три основных свойства `values`: `formulas`, и `format`. Эти свойства получают или задают значения ячеек, вычисляемые формулы и визуальное форматирование ячеек.

#### <a name="range-sample"></a>Пример диапазона

В следующем примере показано, как создать записи о продажах. В этом сценарии `Range` используются объекты для задания значений, формул и форматов.

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the active worksheet.
  let sheet = context.workbook.worksheets.getActiveWorksheet();

  // Create the headers and format them to stand out.
  let headers = [
    ["Product", "Quantity", "Unit Price", "Totals"]
  ];
  let headerRange = sheet.getRange("B2:E2");
  headerRange.values = headers;
  headerRange.format.fill.color = "#4472C4";
  headerRange.format.font.color = "white";

  // Create the product data rows.
  let productData = [
    ["Almonds", 6, 7.5],
    ["Coffee", 20, 34.5],
    ["Chocolate", 10, 9.56],
  ];
  let dataRange = sheet.getRange("B3:D5");
  dataRange.values = productData;

  // Create the formulas to total the amounts sold.
  let totalFormulas = [
    ["=C3 * D3"],
    ["=C4 * D4"],
    ["=C5 * D5"],
    ["=SUM(E3:E5)"]
  ];
  let totalRange = sheet.getRange("E3:E6");
  totalRange.formulas = totalFormulas;
  totalRange.format.font.bold = true;

  // Display the totals as US dollar amounts.
  totalRange.numberFormat = [["$0.00"]];
}
```

При выполнении этого сценария на текущем листе создаются следующие данные:

![Запись о продажах, в которой показаны строки значений, столбец формул и отформатированные заголовки.](../images/range-sample.png)

### <a name="charts-tables-and-other-data-objects"></a>Диаграммы, таблицы и другие объекты данных

Скрипты могут создавать и управлять структурами данных и зрительными представлениями в Excel. Таблицы и диаграммы — это два наиболее часто используемых объекта, но API поддерживают сводные таблицы, фигуры, изображения и многое другое.

#### <a name="creating-a-table"></a>Создание таблицы

Создание таблиц с использованием диапазонов, заполненных данными. Элементы управления форматированием и таблицами (например, фильтры) автоматически применяются к диапазону.

Следующий сценарий создает таблицу, используя диапазоны из предыдущего примера.

```TypeScript
async function main(context: Excel.RequestContext) {
   let sheet = context.workbook.worksheets.getActiveWorksheet();
   sheet.tables.add("B2:E5", true);
}
```

При выполнении этого скрипта на листе с предыдущими данными создается следующая таблица:

![Таблица, созданная из предыдущей записи о продажах.](../images/table-sample.png)

#### <a name="creating-a-chart"></a>Создание диаграммы

Создание диаграмм для отображения данных в диапазоне. Сценарии позволяют использовать десятки видов диаграмм, каждая из которых может быть изменена в соответствии со своими потребностями.

Следующий сценарий создает простую гистограмму для трех элементов и размещает его на 100 пикселов ниже верхней границы листа.

```TypeScript
async function main(context: Excel.RequestContext) {
  let sheet = context.workbook.worksheets.getActiveWorksheet();
  let chart = sheet.charts.add(Excel.ChartType.columnStacked, sheet.getRange("B3:C5"));
  chart.top = 100;
}
```

При выполнении этого сценария на листе с помощью предыдущей таблицы создается следующая диаграмма:

![Гистограмма, на которой показано количество трех элементов из предыдущей записи о продажах.](../images/chart-sample.png)

### <a name="further-reading-on-the-object-model"></a>Дальнейшие материалы по объектной модели

[Справочная документация по API сценариев Office](/javascript/api/office-scripts/overview) — это полный список объектов, используемых в сценариях Office. Здесь вы можете использовать оглавление, чтобы перейти к любому классу, о котором вы хотите узнать больше. Ниже представлено несколько часто просматриваемых страниц.

- [Chart](/javascript/api/office-scripts/excel/excel.chart)
- [Comment](/javascript/api/office-scripts/excel/excel.comment)
- [PivotTable](/javascript/api/office-scripts/excel/excel.pivottable)
- [Range](/javascript/api/office-scripts/excel/excel.range)
- [RangeFormat](/javascript/api/office-scripts/excel/excel.rangeformat)
- [Shape](/javascript/api/office-scripts/excel/excel.shape)
- [Table](/javascript/api/office-scripts/excel/excel.table)
- [Workbook](/javascript/api/office-scripts/excel/excel.workbook)
- [Worksheet](/javascript/api/office-scripts/excel/excel.worksheet)

## <a name="main-function"></a>`main`функциями

Каждый сценарий Office должен содержать `main` функцию со следующей подписью, в том числе `Excel.RequestContext` определение типа:

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your Excel Script
}
```

Код внутри `main` функции выполняется при запуске скрипта. `main`может вызывать другие функции в вашем сценарии, но код, не содержащийся в функции, не будет выполняться.

## <a name="context"></a>Контекст

`main` Функция принимает `Excel.RequestContext` параметр с именем `context`. Подумайте о `context` качестве моста между сценарием и книгой. Ваш скрипт получает доступ к книге с `context` объектом и использует его `context` для отправки данных назад и вперед.

Этот `context` объект необходим, так как скрипт и Excel выполняются в различных процессах и расположениях. Скрипту потребуется внести изменения в данные или запросить данные из книги в облаке. `context` Объект управляет этими транзакциями.

## <a name="sync-and-load"></a>Синхронизация и загрузка

Так как сценарий и книга работают в различных расположениях, передача данных между ними занимает некоторое время. Для повышения производительности скриптов команды ставятся в очередь до тех пор, пока скрипт не `sync` выполнит явное вызов операции для синхронизации скрипта и книги. Ваш сценарий может работать независимо, пока ему не потребуется выполнить одно из следующих действий:

- Чтение данных из книги (после `load` операции).
- Запись данных в книгу (обычно потому, что сценарий завершен).

На приведенном ниже рисунке показан пример последовательности элементов управления между скриптом и книгой:

![Схема, демонстрирующая операции чтения и записи, идущие к книге из скрипта.](../images/load-sync.png)

### <a name="sync"></a>Синхронизировать

Если сценарий должен считывать данные из книги или записывать данные в нее, вызовите `RequestContext.sync` метод, как показано ниже:

```TypeScript
await context.sync();
```

> [!NOTE]
> `context.sync()`вызывается неявно при завершении скрипта.

После завершения `sync` операции книга обновляется, чтобы отразить все операции записи, заданные в скрипте. Операция записи задает любое свойство объекта Excel (например, `range.format.fill.color = "red"`) или вызывает метод, который изменяет свойство (например, `range.format.autoFitColumns()`). Эта `sync` операция также считывает все значения из книги, запрошенные сценарием с помощью `load` операции (как описано в следующем разделе).

Синхронизация вашего сценария с книгой может занять некоторое время в зависимости от сети. Необходимо минимизировать количество `sync` вызовов, чтобы ускорить выполнение сценария.  

### <a name="load"></a>Load

Перед прочтением скрипту необходимо загрузить данные из книги. Однако часто загрузка данных из всей книги значительно сокращает скорость сценария. Вместо этого `load` метод предоставляет состояние скрипта, в частности, какие данные следует извлечь из книги.

`load` Метод доступен для каждого объекта Excel. Скрипту необходимо загрузить свойства объекта, прежде чем он сможет прочитать их. В противном случае будет возникать ошибка.

В приведенных ниже примерах используется `Range` объект для отображения трех способов `load` , которые метод может использовать для загрузки данных.

|Условие |Пример команды | Эффект |
|:--|:--|:--|
|Загрузить одно свойство |`myRange.load("values");` | Загружает одно свойство, в данном случае двухмерный массив значений из этого диапазона. |
|Загрузка нескольких свойств |`myRange.load("values, rowCount, columnCount");`| Загружает все свойства из списка, разделенного запятыми, в данном примере значения, количество строк и число столбцов. |
|Загрузка всех элементов | `myRange.load();`|Загружает все свойства диапазона. Это не рекомендуемое решение, так как оно замедляет выполнение скрипта, получая ненужные данные. Этот параметр следует использовать только при тестировании скрипта или при необходимости каждого свойства объекта. |

Перед чтением любых загруженных значений ваш сценарий должен вызываться `context.sync()` .

```TypeScript
let range = selectedSheet.getRange("A1:B3");
range.load ("rowCount"); // Load the property.
await context.sync(); // Synchronize with the workbook to get the property.
console.log(range.rowCount); // Read and log the property value (3).
```

Кроме того, можно загружать свойства во всей коллекции. У каждого объекта Collection есть `items` свойство, которое представляет собой массив, содержащий объекты в этой коллекции. Использование `items` в качестве начала иерархического вызова (`items\myProperty`) для `load` загрузки указанных свойств для каждого из этих элементов. В следующем примере загружается `resolved` свойство для `Comment` каждого объекта в `CommentCollection` объекте листа.

```TypeScript
let comments = selectedSheet.comments;
comments.load("items/resolved"); // Load the `resolved` property from every comment in this collection.
await context.sync(); // Synchronize with the workbook to get the properties.
```

> [!TIP]
> Чтобы узнать больше о работе с коллекциями в сценариях Office, ознакомьтесь с [разделом Array, посвященным встроенным объектам JavaScript в статье сценариев Office](javascript-objects.md#array) .

## <a name="see-also"></a>См. также

- [Запись, редактирование и создание сценариев Office в Excel в Интернете](../tutorials/excel-tutorial.md)
- [Чтение данных книги с помощью скриптов Office в Excel в Интернете](../tutorials/excel-read-tutorial.md)
- [Справочник по API скриптов Office](/javascript/api/office-scripts/overview)
- [Использование встроенных объектов JavaScript в сценариях Office](javascript-objects.md)
