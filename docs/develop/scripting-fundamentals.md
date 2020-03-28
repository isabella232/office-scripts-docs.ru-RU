---
title: Основы сценариев для сценариев Office в Excel в Интернете
description: Информация об объектной модели и другие основы для изучения перед написанием сценариев Office.
ms.date: 01/27/2020
localization_priority: Priority
ms.openlocfilehash: 5a709c16e23c00ffc7ee7949a3cb11459dc2d530
ms.sourcegitcommit: d556aaefac80e55f53ac56b7f6ecbc657ebd426f
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2020
ms.locfileid: "42978734"
---
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web-preview"></a>Основы сценариев для сценариев Office в Excel в Интернете (предварительная версия)

Эта статья познакомит вас с техническими аспектами сценариев Office. Вы узнаете, как объекты Excel работают вместе и как редактор кода синхронизируется с книгой.

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="object-model"></a>Объектная модель

Чтобы понять API-интерфейсы Excel, вы должны понимать, как компоненты рабочей книги связаны друг с другом.

- **Рабочая книга** содержит одну или несколько **рабочих листов**.
- **Рабочий лист** предоставляет доступ к ячейкам через объекты **Range**.
- **Range** представляет группу смежных клеток.
- **Диапазоны** используются для создания и размещения **таблиц**, **диаграмм**, **фигур** и других объектов визуализации данных или организации.
- **Рабочий лист** содержит коллекции тех объектов данных, которые присутствуют на отдельном листе.
- **Рабочие книги** содержат коллекции некоторых из этих объектов данных (таких как **таблицы**) для всей **рабочей книги**.

### <a name="ranges"></a>Диапазоны

Диапазон - это группа непрерывных ячеек в рабочей книге. В сценариях обычно используется нотация в стиле A1 (например, **B3** для отдельной ячейки в строке **B** и столбце **3** или **C2:F4** для ячеек из строк с **C** по **F** и столбцов со **2** по **4**) для определения диапазонов.

Диапазоны имеют три основных свойства: `values`, `formulas`, и `format`. Эти свойства получают или устанавливают значения ячеек, формулы для оценки и визуальное форматирование ячеек.

#### <a name="range-sample"></a>Образец диапазона

В следующем примере показано, как создавать записи продаж. Этот скрипт использует `Range` объекты для установки значений, формул и форматов.

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

Выполнение этого скрипта создает следующие данные в текущей рабочей таблице:

![Запись о продажах, показывающая строки значений, столбец формулы и отформатированные заголовки.](../images/range-sample.png)

### <a name="charts-tables-and-other-data-objects"></a>Диаграммы, таблицы и другие объекты данных

Скрипты могут создавать и управлять структурами данных и визуализациями в Excel. Таблицы и диаграммы являются двумя наиболее часто используемыми объектами, но API поддерживают сводные таблицы, фигуры, изображения и многое другое.

#### <a name="creating-a-table"></a>Создание таблицы

Создавайте таблицы, используя заполненные данными диапазоны. Элементы управления форматированием и таблицами (например, фильтры) автоматически применяются к диапазону.

Следующий скрипт создает таблицу с использованием диапазонов из предыдущего примера.

```TypeScript
async function main(context: Excel.RequestContext) {
   let sheet = context.workbook.worksheets.getActiveWorksheet();
   sheet.tables.add("B2:E5", true);
}
```

Выполнение этого сценария на листе с предыдущими данными создает следующую таблицу:

![Таблица сделана из предыдущего рекорда продаж.](../images/table-sample.png)

#### <a name="creating-a-chart"></a>Создание диаграммы

Создайте диаграммы для визуализации данных в диапазоне. Сценарии позволяют создавать десятки разновидностей диаграмм, каждая из которых может быть настроена в соответствии с вашими потребностями.

Следующий скрипт создает простую столбчатую диаграмму для трех элементов и размещает ее на 100 пикселей ниже верхней части листа.

```TypeScript
async function main(context: Excel.RequestContext) {
  let sheet = context.workbook.worksheets.getActiveWorksheet();
  let chart = sheet.charts.add(Excel.ChartType.columnStacked, sheet.getRange("B3:C5"));
  chart.top = 100;
}
```

Запуск этого скрипта на листе с предыдущей таблицей создает следующую диаграмму:

![Столбчатая диаграмма, показывающая количество трех предметов из предыдущей записи о продажах.](../images/chart-sample.png)

### <a name="further-reading-on-the-object-model"></a>Дальнейшее чтение по объектной модели

[Справочная документация по API сценариев Office](/javascript/api/office-scripts/overview) представляет собой полный список объектов, используемых в сценариях Office. Там вы можете использовать оглавление, чтобы перейти к любому классу, о котором вы хотите узнать больше. Ниже приведены несколько часто просматриваемых страниц.

- [Chart](/javascript/api/office-scripts/excel/excel.chart)
- [Comment](/javascript/api/office-scripts/excel/excel.comment)
- [PivotTable](/javascript/api/office-scripts/excel/excel.pivottable)
- [Range](/javascript/api/office-scripts/excel/excel.range)
- [RangeFormat](/javascript/api/office-scripts/excel/excel.rangeformat)
- [Shape](/javascript/api/office-scripts/excel/excel.shape)
- [Table](/javascript/api/office-scripts/excel/excel.table)
- [Workbook](/javascript/api/office-scripts/excel/excel.workbook)
- [Worksheet](/javascript/api/office-scripts/excel/excel.worksheet)

## <a name="main-function"></a>`main` функция

Каждый сценарий Office должен содержать `main` функцию со следующей подписью, включая определение `Excel.RequestContext` типа:

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your Excel Script
}
```

Код внутри `main` функции запускается при запуске скрипта. `main` может вызывать другие функции в вашем скрипте, но код, который не содержится в функции, не будет работать.

## <a name="context"></a>Context

Функция `main` принимает `Excel.RequestContext` параметра с именем `context`. Думайте о `context` как о мосте между вашим сценарием и книгой. Ваш сценарий обращается к книге с помощью `context` объекта и использует этот `context` для отправки данных туда и обратно.

Объект `context` необходим, потому что скрипт и Excel работают в разных процессах и местах. Сценарий должен будет внести изменения или запросить данные из рабочей книги в облаке. Объект `context` управляет этими транзакциями.

## <a name="sync-and-load"></a>Синхронизация и загрузка

Поскольку ваш сценарий и рабочая книга работают в разных местах, любая передача данных между ними занимает много времени. Для повышения производительности сценария команды помещаются в очередь до тех пор, пока сценарий явно не вызовет `sync` операцию для синхронизации сценария и рабочей книги. Ваш скрипт может работать независимо, пока он не выполнит одно из следующих действий:

- Чтение данных из рабочей книги (после `load` операции).
- Запишите данные в рабочую книгу (обычно потому, что сценарий завершен).

На следующем рисунке показан пример потока управления между сценарием и книгой:

![Диаграмма, показывающая операции чтения и записи, идущие в рабочую книгу из сценария.](../images/load-sync.png)

### <a name="sync"></a>Синхронизировать

Всякий раз, когда вашему сценарию нужно прочитать данные или записать данные в рабочую книгу, вызывайте метод `RequestContext.sync`, как показано здесь:

```TypeScript
await context.sync();
```

> [!NOTE]
> `context.sync()` неявно вызывается, когда скрипт заканчивается.

После завершения операции `sync` книга обновляется, чтобы отразить все операции записи, указанные сценарием. Операция записи устанавливает любое свойство для объекта Excel (например, `range.format.fill.color = "red"`) или вызывает метод, который изменяет свойство (например, `range.format.autoFitColumns()`). Операция `sync` также считывает любые значения из рабочей книги, запрошенные сценарием с помощью операции `load` (как описано в следующем разделе).

Синхронизация вашего сценария с книгой может занять некоторое время, в зависимости от вашей сети. Вы должны минимизировать количество вызовов `sync`, чтобы ваш скрипт работал быстро.  

### <a name="load"></a>Load

Сценарий должен загрузить данные из рабочей книги перед ее чтением. Однако частая загрузка данных из всей рабочей книги значительно снижает скорость работы сценария. Вместо этого метод `load` позволяет вашему сценарию указать, какие именно данные следует извлечь из рабочей книги.

Метод `load` доступен для каждого объекта Excel. Ваш скрипт должен загрузить свойства объекта, прежде чем он сможет их прочитать. Невыполнение этого требования приведет к ошибке.

В следующих примерах объект `Range` используется для демонстрации трех способов использования метода `load` для загрузки данных.

|Intent |Пример команды | Эффект |
|:--|:--|:--|
|Загрузить одно свойство |`myRange.load("values");` | Загружает одно свойство, в данном случае двумерный массив значений в этом диапазоне. |
|Загрузить несколько свойств |`myRange.load("values, rowCount, columnCount");`| Загружает все свойства из списка, разделенного запятыми, в этом примере значения, количество строк и количество столбцов. |
|Загрузить все | `myRange.load();`|Загружает все свойства в диапазоне. Это не рекомендуемое решение, так как оно замедлит ваш скрипт, получая ненужные данные. Вы должны использовать это только при тестировании вашего скрипта или если вам нужно каждое свойство объекта. |

Ваш скрипт должен вызывать `context.sync()` перед чтением любых загруженных значений.

```TypeScript
let range = selectedSheet.getRange("A1:B3");
range.load ("rowCount"); // Load the property.
await context.sync(); // Synchronize with the workbook to get the property.
console.log(range.rowCount); // Read and log the property value (3).
```

Вы также можете загрузить свойства всей коллекции. Каждый объект коллекции имеет `items` свойство, которое является массивом, содержащим объекты в этой коллекции. Использование `items` в качестве начала иерархического вызова (`items\myProperty`) для `load` загружает указанные свойства для каждого из этих элементов. В следующем примере загружается свойство `resolved` для каждых `Comment` объектов в `CommentCollection` объекте рабочего листа.

```TypeScript
let comments = selectedSheet.comments;
comments.load("items/resolved"); // Load the `resolved` property from every comment in this collection.
await context.sync(); // Synchronize with the workbook to get the properties.
```

> [!TIP]
> Подробнее о работе с коллекциями в сценариях Office см. в статье, в разделе [массива Использование встроенных объектов JavaScript в сценариях Office](javascript-objects.md#array).

## <a name="see-also"></a>См. также

- [Запись, редактирование и создание сценариев Office в Excel в Интернете](../tutorials/excel-tutorial.md)
- [Чтение данных рабочей книги с помощью сценариев Office в Excel в Интернете](../tutorials/excel-read-tutorial.md)
- [Справочник API для сценариев Office](/javascript/api/office-scripts/overview)
- [Использование встроенных объектов JavaScript в сценариях Office](javascript-objects.md)
