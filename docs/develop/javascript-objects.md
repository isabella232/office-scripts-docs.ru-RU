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
# <a name="using-built-in-javascript-objects-in-office-scripts"></a><span data-ttu-id="f9287-103">Использование встроенных объектов JavaScript в сценариях Office</span><span class="sxs-lookup"><span data-stu-id="f9287-103">Using built-in JavaScript objects in Office Scripts</span></span>

<span data-ttu-id="f9287-104">JavaScript предоставляет несколько встроенных объектов, которые можно использовать в сценариях Office независимо от того, используете ли вы скрипты в JavaScript или [TypeScript](../overview/code-editor-environment.md) (расширенный набор JavaScript).</span><span class="sxs-lookup"><span data-stu-id="f9287-104">JavaScript provides several built-in objects that you can use in your Office Scripts, regardless of whether you're scripting in JavaScript or [TypeScript](../overview/code-editor-environment.md) (a superset of JavaScript).</span></span> <span data-ttu-id="f9287-105">В этой статье описывается, как можно использовать встроенные объекты JavaScript в сценариях Office для Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="f9287-105">This article describes how you can use some of the built-in JavaScript objects in Office Scripts for Excel on the web.</span></span>

> [!NOTE]
> <span data-ttu-id="f9287-106">Полный список всех встроенных объектов JavaScript представлен в статье [стандартные встроенные объекты](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) Mozilla Mozilla.</span><span class="sxs-lookup"><span data-stu-id="f9287-106">For a complete list of all built-in JavaScript objects, see Mozilla's [Standard built-in objects](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) article.</span></span>

## <a name="array"></a><span data-ttu-id="f9287-107">Массив</span><span class="sxs-lookup"><span data-stu-id="f9287-107">Array</span></span>

<span data-ttu-id="f9287-108">Объект [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) обеспечивает стандартизированный способ работы с массивами в вашем скрипте.</span><span class="sxs-lookup"><span data-stu-id="f9287-108">The [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) object provides a standardized way to work with arrays in your script.</span></span> <span data-ttu-id="f9287-109">Хотя массивы являются стандартными конструкциями JavaScript, они относятся к сценариям Office двумя основными способами: Ranges и Collections.</span><span class="sxs-lookup"><span data-stu-id="f9287-109">While arrays are standard JavaScript constructs, they relate to Office Scripts in two major ways: ranges and collections.</span></span>

### <a name="working-with-ranges"></a><span data-ttu-id="f9287-110">Работа с диапазонами</span><span class="sxs-lookup"><span data-stu-id="f9287-110">Working with ranges</span></span>

<span data-ttu-id="f9287-111">Диапазоны содержат несколько двухмерных массивов, которые напрямую сопоставляются с ячейками в этом диапазоне.</span><span class="sxs-lookup"><span data-stu-id="f9287-111">Ranges contain several two-dimensional arrays that directly map to the cells in that range.</span></span> <span data-ttu-id="f9287-112">К ним относятся такие свойства `values`, `formulas`как, `numberFormat`и.</span><span class="sxs-lookup"><span data-stu-id="f9287-112">These include properties such as `values`, `formulas`, and `numberFormat`.</span></span> <span data-ttu-id="f9287-113">Свойства типа array должны [загружаться](scripting-fundamentals.md#sync-and-load) так же, как и любые другие свойства.</span><span class="sxs-lookup"><span data-stu-id="f9287-113">Array-type properties must be [loaded](scripting-fundamentals.md#sync-and-load) like any other properties.</span></span>

<span data-ttu-id="f9287-114">Следующий сценарий выполняет поиск любого числового формата в диапазоне **a1: D4** для любого числового формата, содержащего "$".</span><span class="sxs-lookup"><span data-stu-id="f9287-114">The following script searches the **A1:D4** range for any number format containing a "$".</span></span> <span data-ttu-id="f9287-115">В этом сценарии для цвета заливки в ячейках задается значение "Yellow".</span><span class="sxs-lookup"><span data-stu-id="f9287-115">The script sets the fill color in those cells to "yellow".</span></span>

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

### <a name="working-with-collections"></a><span data-ttu-id="f9287-116">Работа с коллекциями</span><span class="sxs-lookup"><span data-stu-id="f9287-116">Working with collections</span></span>

<span data-ttu-id="f9287-117">В коллекции присутствует множество объектов Excel.</span><span class="sxs-lookup"><span data-stu-id="f9287-117">Many Excel objects are contained in a collection.</span></span> <span data-ttu-id="f9287-118">Например, все [фигуры](/javascript/api/office-scripts/excel/excel.shape) на листе включены в [ShapeCollection](/javascript/api/office-scripts/excel/excel.shapecollection) (как `Worksheet.shapes` свойство).</span><span class="sxs-lookup"><span data-stu-id="f9287-118">For example, all [Shapes](/javascript/api/office-scripts/excel/excel.shape) in a worksheet are contained in a [ShapeCollection](/javascript/api/office-scripts/excel/excel.shapecollection) (as the `Worksheet.shapes` property).</span></span> <span data-ttu-id="f9287-119">Каждый `*Collection` объект содержит `items` свойство, представляющее собой массив, в котором хранятся объекты в этой коллекции.</span><span class="sxs-lookup"><span data-stu-id="f9287-119">Each `*Collection` object contains an `items` property, which is an array that stores the objects inside that collection.</span></span> <span data-ttu-id="f9287-120">Это можно рассматривать как обычный массив JavaScript, но сначала необходимо загрузить элементы коллекции.</span><span class="sxs-lookup"><span data-stu-id="f9287-120">This can be treated like a normal JavaScript array, but the items in the collection have to first be loaded.</span></span> <span data-ttu-id="f9287-121">Если необходимо работать со свойством для каждого объекта в коллекции, используйте инструкцию-иерархию Load (`items/propertyName`).</span><span class="sxs-lookup"><span data-stu-id="f9287-121">If you need to work with a property on every object in the collection, use a hierarchal load statement (`items/propertyName`).</span></span>

<span data-ttu-id="f9287-122">Следующий сценарий записывает в журнал тип каждой фигуры на текущем листе.</span><span class="sxs-lookup"><span data-stu-id="f9287-122">The following script logs the type of every shape in the current worksheet.</span></span>

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

<span data-ttu-id="f9287-123">Можно загружать отдельные объекты из коллекции с помощью методов `getItem` или. `getItemAt`</span><span class="sxs-lookup"><span data-stu-id="f9287-123">You can load individual objects from a collection using the `getItem` or `getItemAt` methods.</span></span> <span data-ttu-id="f9287-124">`getItem`Получает объект с помощью уникального идентификатора, такого как имя (такие имена часто задаются сценарием).</span><span class="sxs-lookup"><span data-stu-id="f9287-124">`getItem` gets an object by using a unique identifier like a name (such names are often specified by your script).</span></span> <span data-ttu-id="f9287-125">`getItemAt`Получает объект, используя его индекс в коллекции.</span><span class="sxs-lookup"><span data-stu-id="f9287-125">`getItemAt` gets an object by using its index in the collection.</span></span> <span data-ttu-id="f9287-126">Прежде чем использовать объект, перед вызовом необходимо указать `await context.sync();` команду.</span><span class="sxs-lookup"><span data-stu-id="f9287-126">Either call must be followed by a `await context.sync();` command before the object can be used.</span></span>

<span data-ttu-id="f9287-127">Следующий сценарий удаляет самую старую фигуру на текущем листе.</span><span class="sxs-lookup"><span data-stu-id="f9287-127">The following script deletes the oldest shape in the current worksheet.</span></span>

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

## <a name="date"></a><span data-ttu-id="f9287-128">Дата</span><span class="sxs-lookup"><span data-stu-id="f9287-128">Date</span></span>

<span data-ttu-id="f9287-129">Объект [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) предоставляет стандартизированный способ работы с датами в скрипте.</span><span class="sxs-lookup"><span data-stu-id="f9287-129">The [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) object provides a standardized way to work with dates in your script.</span></span> <span data-ttu-id="f9287-130">`Date.now()`Создает объект с текущей датой и временем, который полезен при добавлении меток времени к записи данных в скрипте.</span><span class="sxs-lookup"><span data-stu-id="f9287-130">`Date.now()` generates an object with the current date and time, which is useful when adding timestamps to your script's data entry.</span></span>

<span data-ttu-id="f9287-131">Следующий сценарий добавляет текущую дату на лист.</span><span class="sxs-lookup"><span data-stu-id="f9287-131">The following script adds the current date to the worksheet.</span></span> <span data-ttu-id="f9287-132">Обратите внимание, что `toLocaleDateString` с помощью метода Excel распознает значение как дату и автоматически изменяет формат числа в ячейке.</span><span class="sxs-lookup"><span data-stu-id="f9287-132">Note that by using the `toLocaleDateString` method, Excel recognizes the value as a date and changes the number format of the cell automatically.</span></span>

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

<span data-ttu-id="f9287-133">В разделе " [трудозатраты с датами](../resources/excel-samples.md#work-with-dates) " в примерах имеются дополнительные скрипты, связанные с датами.</span><span class="sxs-lookup"><span data-stu-id="f9287-133">The [Work with dates](../resources/excel-samples.md#work-with-dates) section of the samples has more Date-related scripts.</span></span>

## <a name="math"></a><span data-ttu-id="f9287-134">математика;</span><span class="sxs-lookup"><span data-stu-id="f9287-134">Math</span></span>

<span data-ttu-id="f9287-135">Объект [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) предоставляет методы и константы для распространенных математических операций.</span><span class="sxs-lookup"><span data-stu-id="f9287-135">The [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) object provides methods and constants for common mathematical operations.</span></span> <span data-ttu-id="f9287-136">Они предоставляют множество функций, которые также доступны в Excel, без необходимости использовать подсистему вычисления книги.</span><span class="sxs-lookup"><span data-stu-id="f9287-136">These provide many functions also available in Excel, without the need to use the workbook's calculation engine.</span></span> <span data-ttu-id="f9287-137">При этом скрипту не требуется запрашивать книгу, что повышает производительность.</span><span class="sxs-lookup"><span data-stu-id="f9287-137">This saves your script from having to query the workbook, which improves performance.</span></span>

<span data-ttu-id="f9287-138">Следующий сценарий использует `Math.min` для поиска и записи в журнал наименьшего числа в диапазоне **a1: D4** .</span><span class="sxs-lookup"><span data-stu-id="f9287-138">The following script uses `Math.min` to find and log the smallest number in the **A1:D4** range.</span></span> <span data-ttu-id="f9287-139">Обратите внимание, что в этом примере предполагается, что весь диапазон содержит только цифры, а не строки.</span><span class="sxs-lookup"><span data-stu-id="f9287-139">Note that this sample assumes the entire range contains only numbers, not strings.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="f9287-140">См. также</span><span class="sxs-lookup"><span data-stu-id="f9287-140">See also</span></span>

- [<span data-ttu-id="f9287-141">Стандартные встроенные объекты</span><span class="sxs-lookup"><span data-stu-id="f9287-141">Standard built-in objects</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [<span data-ttu-id="f9287-142">Среда редактора кода сценариев Office</span><span class="sxs-lookup"><span data-stu-id="f9287-142">Office Scripts Code Editor environment</span></span>](../overview/code-editor-environment.md)
