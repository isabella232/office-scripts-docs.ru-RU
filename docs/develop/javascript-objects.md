---
title: Использование встроенных объектов JavaScript в сценариях Office
description: Как вызывать встроенные API JavaScript из скрипта Office в Excel в Интернете.
ms.date: 04/06/2020
localization_priority: Normal
ms.openlocfilehash: a4b698215edea5f266e159fee0e08690904dd379
ms.sourcegitcommit: 0b2232c4c228b14d501edb8bb489fe0e84748b42
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/08/2020
ms.locfileid: "43191017"
---
# <a name="using-built-in-javascript-objects-in-office-scripts"></a>Использование встроенных объектов JavaScript в сценариях Office

JavaScript предоставляет несколько встроенных объектов, которые можно использовать в сценариях Office независимо от того, используете ли вы скрипты в JavaScript или [TypeScript](../overview/code-editor-environment.md) (расширенный набор JavaScript). В этой статье описывается, как можно использовать встроенные объекты JavaScript в сценариях Office для Excel в Интернете.

> [!NOTE]
> Полный список всех встроенных объектов JavaScript представлен в статье [стандартные встроенные объекты](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) Mozilla Mozilla.

## <a name="array"></a>Массив

Объект [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) обеспечивает стандартизированный способ работы с массивами в вашем скрипте. Хотя массивы являются стандартными конструкциями JavaScript, они относятся к сценариям Office двумя основными способами: Ranges и Collections.

### <a name="working-with-ranges"></a>Работа с диапазонами

Диапазоны содержат несколько двухмерных массивов, которые напрямую сопоставляются с ячейками в этом диапазоне. К ним относятся такие свойства `values`, `formulas`как, `numberFormat`и. Свойства типа array должны [загружаться](scripting-fundamentals.md#sync-and-load) так же, как и любые другие свойства.

Следующий сценарий выполняет поиск любого числового формата в диапазоне **a1: D4** для любого числового формата, содержащего "$". В этом сценарии для цвета заливки в ячейках задается значение "Yellow".

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the range From A1 to D4.
  let range = context.workbook.worksheets.getActiveWorksheet().getRange("A1:D4");

  // Load the numberFormat property on the range.
  range.load("numberFormat");
  await context.sync();

  // Iterate through the arrays of rows and columns corresponding to those in the range.
  range.numberFormat.forEach((rowItem, rowIndex) => {
    range.numberFormat[rowIndex].forEach((columnItem, columnIndex) => {
      // Treat the numberFormat as a string so we can do text comparisons.
      let columnItemText = columnItem as string;
      if (columnItemText.indexOf("$") >= 0) {
        // Set the cell's fill to yellow.
        range.getCell(rowIndex, columnIndex).format.fill.color = "yellow";
      }
    });
  });
}
```

### <a name="working-with-collections"></a>Работа с коллекциями

В коллекции присутствует множество объектов Excel. Например, все [фигуры](/javascript/api/office-scripts/excel/excel.shape) на листе включены в [ShapeCollection](/javascript/api/office-scripts/excel/excel.shapecollection) (как `Worksheet.shapes` свойство). Каждый `*Collection` объект содержит `items` свойство, представляющее собой массив, в котором хранятся объекты в этой коллекции. Это можно рассматривать как обычный массив JavaScript, но сначала необходимо загрузить элементы коллекции. Если необходимо работать со свойством для каждого объекта в коллекции, используйте инструкцию-иерархию Load (`items/propertyName`).

Следующий сценарий записывает в журнал тип каждой фигуры на текущем листе.

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Get the shapes in this worksheet.
  let shapes = selectedSheet.shapes;
  shapes.load("items/type");
  await context.sync();

  // Log the type of every shape in the collection.
  shapes.items.forEach((shape) => {
    console.log(shape.type);
  });
}
```

Можно загружать отдельные объекты из коллекции с помощью методов `getItem` или. `getItemAt` `getItem`Получает объект с помощью уникального идентификатора, такого как имя (такие имена часто задаются сценарием). `getItemAt`Получает объект, используя его индекс в коллекции. Прежде чем использовать объект, перед вызовом необходимо указать `await context.sync();` команду.

Следующий сценарий удаляет самую старую фигуру на текущем листе.

```Typescript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Get the first (oldest) shape in the worksheet.
  // Note that this script will thrown an error if there are no shapes.
  let shape = selectedSheet.shapes.getItemAt(0);

  // Sync to load `shape` from the collection.
  await context.sync();

  // Remove the shape from the worksheet.
  shape.delete();
}
```

## <a name="date"></a>Дата

Объект [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) предоставляет стандартизированный способ работы с датами в скрипте. `Date.now()`Создает объект с текущей датой и временем, который полезен при добавлении меток времени к записи данных в скрипте.

Следующий сценарий добавляет текущую дату на лист. Обратите внимание, что `toLocaleDateString` с помощью метода Excel распознает значение как дату и автоматически изменяет формат числа в ячейке.

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the range for cell A1.
  let range = context.workbook.worksheets.getActiveWorksheet().getRange("A1");

  // Get the current date and time.
  let date = new Date(Date.now());

  // Set the value at A1 to the current date, using a localized string.
  range.values = [[date.toLocaleDateString()]];
}
```

В разделе " [трудозатраты с датами](../resources/excel-samples.md#work-with-dates) " в примерах имеются дополнительные скрипты, связанные с датами.

## <a name="math"></a>математика;

Объект [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) предоставляет методы и константы для распространенных математических операций. Они предоставляют множество функций, которые также доступны в Excel, без необходимости использовать подсистему вычисления книги. При этом скрипту не требуется запрашивать книгу, что повышает производительность.

Следующий сценарий использует `Math.min` для поиска и записи в журнал наименьшего числа в диапазоне **a1: D4** . Обратите внимание, что в этом примере предполагается, что весь диапазон содержит только цифры, а не строки.

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the range from A1 to D4.
  let comparisonRange = context.workbook.worksheets.getActiveWorksheet().getRange("A1:D4");
  
  // Load the range's values.
  comparisonRange.load("values");
  await context.sync();

  // Set the minimum values as the first value.
  let minimum = comparisonRange.values[0][0];

  // Iterate over each row looking for the smallest value.
  comparisonRange.values.forEach((rowItem, rowIndex) => {
    // Iterate over each column looking for the smallest value.
    comparisonRange.values[rowIndex].forEach((columnItem) => {
      // Use `Math.min` to set the smallest value as either the current cell's value or the previous minimum.
      minimum = Math.min(minimum, columnItem);
    });
  });
  
  console.log(minimum);
}

```

## <a name="see-also"></a>См. также

- [Стандартные встроенные объекты](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [Среда редактора кода сценариев Office](../overview/code-editor-environment.md)
