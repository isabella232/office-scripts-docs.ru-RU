---
title: Использование встроенных объектов JavaScript в сценариях Office
description: Как вызывать встроенные API JavaScript из скрипта Office в Excel в Интернете.
ms.date: 07/16/2020
localization_priority: Normal
ms.openlocfilehash: 4bb5fb5444887005ececbbfdf0130cba3784e0c4
ms.sourcegitcommit: 8d549884e68170f808d3d417104a4451a37da83c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/22/2020
ms.locfileid: "45229598"
---
# <a name="using-built-in-javascript-objects-in-office-scripts"></a>Использование встроенных объектов JavaScript в сценариях Office

JavaScript предоставляет несколько встроенных объектов, которые можно использовать в сценариях Office независимо от того, используете ли вы скрипты в JavaScript или [TypeScript](../overview/code-editor-environment.md) (расширенный набор JavaScript). В этой статье описывается, как можно использовать встроенные объекты JavaScript в сценариях Office для Excel в Интернете.

> [!NOTE]
> Полный список всех встроенных объектов JavaScript представлен в статье [стандартные встроенные объекты](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) Mozilla Mozilla.

## <a name="array"></a>Массив

Объект [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) обеспечивает стандартизированный способ работы с массивами в вашем скрипте. Хотя массивы являются стандартными конструкциями JavaScript, они относятся к сценариям Office двумя основными способами: Ranges и Collections.

### <a name="working-with-ranges"></a>Работа с диапазонами

Диапазоны содержат несколько двухмерных массивов, которые напрямую сопоставляются с ячейками в этом диапазоне. Эти массивы содержат конкретные сведения о каждой ячейке в этом диапазоне. Например, `Range.getValues` возвращает все значения в этих ячейках (со строками и столбцами, которые сопоставлены с двумерным массивом, на строки и столбцы этого подраздела листа). `Range.getFormulas`и `Range.getNumberFormats` это часто используемые методы, возвращающие массивы, такие как `Range.getValues` .

Следующий сценарий выполняет поиск любого числового формата в диапазоне **a1: D4** для любого числового формата, содержащего "$". В этом сценарии для цвета заливки в ячейках задается значение "Yellow".

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the range From A1 to D4.
  let range = workbook.getActiveWorksheet().getRange("A1:D4");

  // Get the number formats for each cell in the range.
  let rangeNumberFormats = range.getNumberFormats();
  // Iterate through the arrays of rows and columns corresponding to those in the range.
  rangeNumberFormats.forEach((rowItem, rowIndex) => {
    rangeNumberFormats[rowIndex].forEach((columnItem, columnIndex) => {
      // Treat the numberFormat as a string so we can do text comparisons.
      let columnItemText = columnItem as string;
      if (columnItemText.indexOf("$") >= 0) {
        // Set the cell's fill to yellow.
        range.getCell(rowIndex, columnIndex).getFormat().getFill().setColor("yellow");
      }
    });
  });
}
```

### <a name="working-with-collections"></a>Работа с коллекциями

В коллекции присутствует множество объектов Excel. Коллекция управляется API скриптов Office и предоставляется в виде массива. Например, все [фигуры](/javascript/api/office-scripts/excelscript/excelscript.shape) на листе включены в объект `Shape[]` , возвращаемый `Worksheet.getShapes` методом. Этот массив можно использовать для считывания значений из коллекции или для доступа к определенным объектам из методов родительского объекта `get*` .

> [!NOTE]
> Не добавляйте и не удаляйте объекты из этих массивов коллекций вручную. Используйте `add` методы для родительских объектов и `delete` методы в объектах типа Collection. Например, добавьте [таблицу](/javascript/api/office-scripts/excelscript/excelscript.table) на [лист](/javascript/api/office-scripts/excelscript/excelscript.worksheet) с `Worksheet.addTable` методом и удалите метод `Table` using `Table.delete` .

Следующий сценарий записывает в журнал тип каждой фигуры на текущем листе.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the shapes in this worksheet.
  let shapes = selectedSheet.getShapes();

  // Log the type of every shape in the collection.
  shapes.forEach((shape) => {
    console.log(shape.getType());
  });
}
```

Следующий сценарий удаляет самую старую фигуру на текущем листе.

```Typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the first (oldest) shape in the worksheet.
  // Note that this script will thrown an error if there are no shapes.
  let shape = selectedSheet.getShapes()[0];

  // Remove the shape from the worksheet.
  shape.delete();
}
```

## <a name="date"></a>Дата

Объект [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) предоставляет стандартизированный способ работы с датами в скрипте. `Date.now()`Создает объект с текущей датой и временем, который полезен при добавлении меток времени к записи данных в скрипте.

Следующий сценарий добавляет текущую дату на лист. Обратите внимание, что с помощью `toLocaleDateString` метода Excel распознает значение как дату и автоматически изменяет формат числа в ячейке.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the range for cell A1.
  let range = workbook.getActiveWorksheet().getRange("A1");

  // Get the current date and time.
  let date = new Date(Date.now());

  // Set the value at A1 to the current date, using a localized string.
  range.setValue(date.toLocaleDateString());
}
```

В разделе " [трудозатраты с датами](../resources/excel-samples.md#dates) " в примерах имеются дополнительные скрипты, связанные с датами.

## <a name="math"></a>математика;

Объект [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) предоставляет методы и константы для распространенных математических операций. Они предоставляют множество функций, которые также доступны в Excel, без необходимости использовать подсистему вычисления книги. При этом скрипту не требуется запрашивать книгу, что повышает производительность.

Следующий сценарий использует `Math.min` для поиска и записи в журнал наименьшего числа в диапазоне **a1: D4** . Обратите внимание, что в этом примере предполагается, что весь диапазон содержит только цифры, а не строки.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the range from A1 to D4.
  let comparisonRange = workbook.getActiveWorksheet().getRange("A1:D4");

  // Load the range's values.
  let comparisonRangeValues = comparisonRange.getValues();

  // Set the minimum values as the first value.
  let minimum = comparisonRangeValues[0][0];

  // Iterate over each row looking for the smallest value.
  comparisonRangeValues.forEach((rowItem, rowIndex) => {
    // Iterate over each column looking for the smallest value.
    comparisonRangeValues[rowIndex].forEach((columnItem) => {
      // Use `Math.min` to set the smallest value as either the current cell's value or the previous minimum.
      minimum = Math.min(minimum, columnItem);
    });
  });

  console.log(minimum);
}

```

## <a name="use-of-external-javascript-libraries-is-not-supported"></a>Использование внешних библиотек JavaScript не поддерживается

Сценарии Office не поддерживают использование внешних сторонних библиотек. Ваш сценарий может использовать только встроенные объекты JavaScript и API сценариев Office.

## <a name="see-also"></a>См. также

- [Стандартные встроенные объекты](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [Среда редактора кода сценариев Office](../overview/code-editor-environment.md)
