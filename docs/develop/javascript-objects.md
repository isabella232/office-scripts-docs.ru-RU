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
# <a name="using-built-in-javascript-objects-in-office-scripts"></a><span data-ttu-id="664c4-103">Использование встроенных объектов JavaScript в сценариях Office</span><span class="sxs-lookup"><span data-stu-id="664c4-103">Using built-in JavaScript objects in Office Scripts</span></span>

<span data-ttu-id="664c4-104">JavaScript предоставляет несколько встроенных объектов, которые можно использовать в сценариях Office независимо от того, используете ли вы скрипты в JavaScript или [TypeScript](../overview/code-editor-environment.md) (расширенный набор JavaScript).</span><span class="sxs-lookup"><span data-stu-id="664c4-104">JavaScript provides several built-in objects that you can use in your Office Scripts, regardless of whether you're scripting in JavaScript or [TypeScript](../overview/code-editor-environment.md) (a superset of JavaScript).</span></span> <span data-ttu-id="664c4-105">В этой статье описывается, как можно использовать встроенные объекты JavaScript в сценариях Office для Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="664c4-105">This article describes how you can use some of the built-in JavaScript objects in Office Scripts for Excel on the web.</span></span>

> [!NOTE]
> <span data-ttu-id="664c4-106">Полный список всех встроенных объектов JavaScript представлен в статье [стандартные встроенные объекты](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) Mozilla Mozilla.</span><span class="sxs-lookup"><span data-stu-id="664c4-106">For a complete list of all built-in JavaScript objects, see Mozilla's [Standard built-in objects](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) article.</span></span>

## <a name="array"></a><span data-ttu-id="664c4-107">Массив</span><span class="sxs-lookup"><span data-stu-id="664c4-107">Array</span></span>

<span data-ttu-id="664c4-108">Объект [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) обеспечивает стандартизированный способ работы с массивами в вашем скрипте.</span><span class="sxs-lookup"><span data-stu-id="664c4-108">The [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) object provides a standardized way to work with arrays in your script.</span></span> <span data-ttu-id="664c4-109">Хотя массивы являются стандартными конструкциями JavaScript, они относятся к сценариям Office двумя основными способами: Ranges и Collections.</span><span class="sxs-lookup"><span data-stu-id="664c4-109">While arrays are standard JavaScript constructs, they relate to Office Scripts in two major ways: ranges and collections.</span></span>

### <a name="working-with-ranges"></a><span data-ttu-id="664c4-110">Работа с диапазонами</span><span class="sxs-lookup"><span data-stu-id="664c4-110">Working with ranges</span></span>

<span data-ttu-id="664c4-111">Диапазоны содержат несколько двухмерных массивов, которые напрямую сопоставляются с ячейками в этом диапазоне.</span><span class="sxs-lookup"><span data-stu-id="664c4-111">Ranges contain several two-dimensional arrays that directly map to the cells in that range.</span></span> <span data-ttu-id="664c4-112">Эти массивы содержат конкретные сведения о каждой ячейке в этом диапазоне.</span><span class="sxs-lookup"><span data-stu-id="664c4-112">These arrays contain specific information about each cell in that range.</span></span> <span data-ttu-id="664c4-113">Например, `Range.getValues` возвращает все значения в этих ячейках (со строками и столбцами, которые сопоставлены с двумерным массивом, на строки и столбцы этого подраздела листа).</span><span class="sxs-lookup"><span data-stu-id="664c4-113">For example, `Range.getValues` returns all the values in those cells (with the rows and columns of the two-dimensional array mapping to the rows and columns of that worksheet subsection).</span></span> <span data-ttu-id="664c4-114">`Range.getFormulas`и `Range.getNumberFormats` это часто используемые методы, возвращающие массивы, такие как `Range.getValues` .</span><span class="sxs-lookup"><span data-stu-id="664c4-114">`Range.getFormulas` and `Range.getNumberFormats` are other frequently used methods that return arrays like `Range.getValues`.</span></span>

<span data-ttu-id="664c4-115">Следующий сценарий выполняет поиск любого числового формата в диапазоне **a1: D4** для любого числового формата, содержащего "$".</span><span class="sxs-lookup"><span data-stu-id="664c4-115">The following script searches the **A1:D4** range for any number format containing a "$".</span></span> <span data-ttu-id="664c4-116">В этом сценарии для цвета заливки в ячейках задается значение "Yellow".</span><span class="sxs-lookup"><span data-stu-id="664c4-116">The script sets the fill color in those cells to "yellow".</span></span>

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

### <a name="working-with-collections"></a><span data-ttu-id="664c4-117">Работа с коллекциями</span><span class="sxs-lookup"><span data-stu-id="664c4-117">Working with collections</span></span>

<span data-ttu-id="664c4-118">В коллекции присутствует множество объектов Excel.</span><span class="sxs-lookup"><span data-stu-id="664c4-118">Many Excel objects are contained in a collection.</span></span> <span data-ttu-id="664c4-119">Коллекция управляется API скриптов Office и предоставляется в виде массива.</span><span class="sxs-lookup"><span data-stu-id="664c4-119">The collection is managed by the Office Scripts API and exposed as an array.</span></span> <span data-ttu-id="664c4-120">Например, все [фигуры](/javascript/api/office-scripts/excelscript/excelscript.shape) на листе включены в объект `Shape[]` , возвращаемый `Worksheet.getShapes` методом.</span><span class="sxs-lookup"><span data-stu-id="664c4-120">For example, all [Shapes](/javascript/api/office-scripts/excelscript/excelscript.shape) in a worksheet are contained in a `Shape[]` that is returned by the `Worksheet.getShapes` method.</span></span> <span data-ttu-id="664c4-121">Этот массив можно использовать для считывания значений из коллекции или для доступа к определенным объектам из методов родительского объекта `get*` .</span><span class="sxs-lookup"><span data-stu-id="664c4-121">You can use this array to read values from the collection, or you can access specific objects from the parent object's `get*` methods.</span></span>

> [!NOTE]
> <span data-ttu-id="664c4-122">Не добавляйте и не удаляйте объекты из этих массивов коллекций вручную.</span><span class="sxs-lookup"><span data-stu-id="664c4-122">Do not manually add or remove objects from these collection arrays.</span></span> <span data-ttu-id="664c4-123">Используйте `add` методы для родительских объектов и `delete` методы в объектах типа Collection.</span><span class="sxs-lookup"><span data-stu-id="664c4-123">Use the `add` methods on the parent objects and the `delete` methods on the collection-type objects.</span></span> <span data-ttu-id="664c4-124">Например, добавьте [таблицу](/javascript/api/office-scripts/excelscript/excelscript.table) на [лист](/javascript/api/office-scripts/excelscript/excelscript.worksheet) с `Worksheet.addTable` методом и удалите метод `Table` using `Table.delete` .</span><span class="sxs-lookup"><span data-stu-id="664c4-124">For example, add a [Table](/javascript/api/office-scripts/excelscript/excelscript.table) to a [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) with the `Worksheet.addTable` method and remove the `Table` using `Table.delete`.</span></span>

<span data-ttu-id="664c4-125">Следующий сценарий записывает в журнал тип каждой фигуры на текущем листе.</span><span class="sxs-lookup"><span data-stu-id="664c4-125">The following script logs the type of every shape in the current worksheet.</span></span>

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

<span data-ttu-id="664c4-126">Следующий сценарий удаляет самую старую фигуру на текущем листе.</span><span class="sxs-lookup"><span data-stu-id="664c4-126">The following script deletes the oldest shape in the current worksheet.</span></span>

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

## <a name="date"></a><span data-ttu-id="664c4-127">Дата</span><span class="sxs-lookup"><span data-stu-id="664c4-127">Date</span></span>

<span data-ttu-id="664c4-128">Объект [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) предоставляет стандартизированный способ работы с датами в скрипте.</span><span class="sxs-lookup"><span data-stu-id="664c4-128">The [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) object provides a standardized way to work with dates in your script.</span></span> <span data-ttu-id="664c4-129">`Date.now()`Создает объект с текущей датой и временем, который полезен при добавлении меток времени к записи данных в скрипте.</span><span class="sxs-lookup"><span data-stu-id="664c4-129">`Date.now()` generates an object with the current date and time, which is useful when adding timestamps to your script's data entry.</span></span>

<span data-ttu-id="664c4-130">Следующий сценарий добавляет текущую дату на лист.</span><span class="sxs-lookup"><span data-stu-id="664c4-130">The following script adds the current date to the worksheet.</span></span> <span data-ttu-id="664c4-131">Обратите внимание, что с помощью `toLocaleDateString` метода Excel распознает значение как дату и автоматически изменяет формат числа в ячейке.</span><span class="sxs-lookup"><span data-stu-id="664c4-131">Note that by using the `toLocaleDateString` method, Excel recognizes the value as a date and changes the number format of the cell automatically.</span></span>

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

<span data-ttu-id="664c4-132">В разделе " [трудозатраты с датами](../resources/excel-samples.md#dates) " в примерах имеются дополнительные скрипты, связанные с датами.</span><span class="sxs-lookup"><span data-stu-id="664c4-132">The [Work with dates](../resources/excel-samples.md#dates) section of the samples has more date-related scripts.</span></span>

## <a name="math"></a><span data-ttu-id="664c4-133">математика;</span><span class="sxs-lookup"><span data-stu-id="664c4-133">Math</span></span>

<span data-ttu-id="664c4-134">Объект [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) предоставляет методы и константы для распространенных математических операций.</span><span class="sxs-lookup"><span data-stu-id="664c4-134">The [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) object provides methods and constants for common mathematical operations.</span></span> <span data-ttu-id="664c4-135">Они предоставляют множество функций, которые также доступны в Excel, без необходимости использовать подсистему вычисления книги.</span><span class="sxs-lookup"><span data-stu-id="664c4-135">These provide many functions also available in Excel, without the need to use the workbook's calculation engine.</span></span> <span data-ttu-id="664c4-136">При этом скрипту не требуется запрашивать книгу, что повышает производительность.</span><span class="sxs-lookup"><span data-stu-id="664c4-136">This saves your script from having to query the workbook, which improves performance.</span></span>

<span data-ttu-id="664c4-137">Следующий сценарий использует `Math.min` для поиска и записи в журнал наименьшего числа в диапазоне **a1: D4** .</span><span class="sxs-lookup"><span data-stu-id="664c4-137">The following script uses `Math.min` to find and log the smallest number in the **A1:D4** range.</span></span> <span data-ttu-id="664c4-138">Обратите внимание, что в этом примере предполагается, что весь диапазон содержит только цифры, а не строки.</span><span class="sxs-lookup"><span data-stu-id="664c4-138">Note that this sample assumes the entire range contains only numbers, not strings.</span></span>

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

## <a name="use-of-external-javascript-libraries-is-not-supported"></a><span data-ttu-id="664c4-139">Использование внешних библиотек JavaScript не поддерживается</span><span class="sxs-lookup"><span data-stu-id="664c4-139">Use of external JavaScript libraries is not supported</span></span>

<span data-ttu-id="664c4-140">Сценарии Office не поддерживают использование внешних сторонних библиотек.</span><span class="sxs-lookup"><span data-stu-id="664c4-140">Office Scripts don't support the use of external, third-party libraries.</span></span> <span data-ttu-id="664c4-141">Ваш сценарий может использовать только встроенные объекты JavaScript и API сценариев Office.</span><span class="sxs-lookup"><span data-stu-id="664c4-141">Your script can only use the built-in JavaScript objects and the Office Scripts APIs.</span></span>

## <a name="see-also"></a><span data-ttu-id="664c4-142">См. также</span><span class="sxs-lookup"><span data-stu-id="664c4-142">See also</span></span>

- [<span data-ttu-id="664c4-143">Стандартные встроенные объекты</span><span class="sxs-lookup"><span data-stu-id="664c4-143">Standard built-in objects</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [<span data-ttu-id="664c4-144">Среда редактора кода сценариев Office</span><span class="sxs-lookup"><span data-stu-id="664c4-144">Office Scripts Code Editor environment</span></span>](../overview/code-editor-environment.md)
