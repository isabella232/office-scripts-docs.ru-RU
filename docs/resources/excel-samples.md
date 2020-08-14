---
title: Примеры сценариев для сценариев Office в Excel в Интернете
description: Коллекция примеров кода для использования со сценариями Office в Excel в Интернете.
ms.date: 08/04/2020
localization_priority: Normal
ms.openlocfilehash: 4f8d6f2395a841a8dcba2ea0e712e645a84a6d91
ms.sourcegitcommit: 1c88abcf5df16a05913f12df89490ce843cfebe2
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/13/2020
ms.locfileid: "46665231"
---
# <a name="sample-scripts-for-office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="0caa3-103">Примеры сценариев для сценариев Office в Excel в Интернете (Предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="0caa3-103">Sample scripts for Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="0caa3-104">Ниже приведены примеры простых сценариев, которые можно использовать в собственных книгах.</span><span class="sxs-lookup"><span data-stu-id="0caa3-104">The following samples are simple scripts for you to try on your own workbooks.</span></span> <span data-ttu-id="0caa3-105">Чтобы использовать их в Excel в Интернете, выполните следующие действия:</span><span class="sxs-lookup"><span data-stu-id="0caa3-105">To use them in Excel on the web:</span></span>

1. <span data-ttu-id="0caa3-106">Откройте вкладку **Автоматизировать**.</span><span class="sxs-lookup"><span data-stu-id="0caa3-106">Open the **Automate** tab.</span></span>
2. <span data-ttu-id="0caa3-107">Нажмите клавишу **Редактор кода**.</span><span class="sxs-lookup"><span data-stu-id="0caa3-107">Press **Code Editor**.</span></span>
3. <span data-ttu-id="0caa3-108">Нажмите **новый скрипт** в области задач редактора кода.</span><span class="sxs-lookup"><span data-stu-id="0caa3-108">Press **New Script** in the Code Editor's task pane.</span></span>
4. <span data-ttu-id="0caa3-109">Замените весь сценарий выбранным образцом.</span><span class="sxs-lookup"><span data-stu-id="0caa3-109">Replace the entire script with the sample of your choice.</span></span>
5. <span data-ttu-id="0caa3-110">В области задач редактора кода нажмите кнопку **запустить** .</span><span class="sxs-lookup"><span data-stu-id="0caa3-110">Press **Run** in the Code Editor's task pane.</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="scripting-basics"></a><span data-ttu-id="0caa3-111">Основные сведения о сценариях</span><span class="sxs-lookup"><span data-stu-id="0caa3-111">Scripting basics</span></span>

<span data-ttu-id="0caa3-112">В этих примерах демонстрируются основные конструктивные блоки для сценариев Office.</span><span class="sxs-lookup"><span data-stu-id="0caa3-112">These samples demonstrate fundamental building blocks for Office Scripts.</span></span> <span data-ttu-id="0caa3-113">Добавьте их в скрипты, чтобы расширить решение и устранить распространенные проблемы.</span><span class="sxs-lookup"><span data-stu-id="0caa3-113">Add these to your scripts to extend your solution and solve common problems.</span></span>

### <a name="read-and-log-one-cell"></a><span data-ttu-id="0caa3-114">Чтение и запись в журнал одной ячейки</span><span class="sxs-lookup"><span data-stu-id="0caa3-114">Read and log one cell</span></span>

<span data-ttu-id="0caa3-115">В этом примере считывается значение **a1** и выводится на консоль.</span><span class="sxs-lookup"><span data-stu-id="0caa3-115">This sample reads the value of **A1** and prints it to the console.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the value of cell A1.
  let range = selectedSheet.getRange("A1");
  
  // Print the value of A1.
  console.log(range.getValue());
}
```

### <a name="read-the-active-cell"></a><span data-ttu-id="0caa3-116">Чтение активной ячейки</span><span class="sxs-lookup"><span data-stu-id="0caa3-116">Read the active cell</span></span>

<span data-ttu-id="0caa3-117">Этот сценарий записывает в журнал значение текущей активной ячейки.</span><span class="sxs-lookup"><span data-stu-id="0caa3-117">This script logs the value of the current active cell.</span></span> <span data-ttu-id="0caa3-118">Если выбрано несколько ячеек, в журнал заносится левая верхняя ячейка.</span><span class="sxs-lookup"><span data-stu-id="0caa3-118">If multiple cells are selected, the top-leftmost cell will be logged.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the current active cell in the workbook.
  let cell = workbook.getActiveCell();

  // Log that cell's value.
  console.log(`The current cell's value is ${cell.getValue()}`);
}
```

### <a name="change-an-adjacent-cell"></a><span data-ttu-id="0caa3-119">Изменение смежной ячейки</span><span class="sxs-lookup"><span data-stu-id="0caa3-119">Change an adjacent cell</span></span>

<span data-ttu-id="0caa3-120">Этот сценарий получает смежные ячейки, используя относительные ссылки.</span><span class="sxs-lookup"><span data-stu-id="0caa3-120">This script gets adjacent cells using relative references.</span></span> <span data-ttu-id="0caa3-121">Обратите внимание, что если активная ячейка находится в верхней строке, часть скрипта завершается с ошибкой, так как она ссылается на ячейку над выбранной в текущий момент.</span><span class="sxs-lookup"><span data-stu-id="0caa3-121">Note that if the active cell is on the top row, part of the script fails, because it references the cell above the currently selected one.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the currently active cell in the workbook.
  let activeCell = workbook.getActiveCell();
  console.log(`The active cell's address is: ${activeCell.getAddress()}`);

  // Get the cell to the right of the active cell and set its value and color.
  let rightCell = activeCell.getOffsetRange(0,1);
  rightCell.setValue("Right cell");
  console.log(`The right cell's address is: ${rightCell.getAddress()}`);
  rightCell.getFormat().getFont().setColor("Magenta");
  rightCell.getFormat().getFill().setColor("Cyan");

  // Get the cell to the above of the active cell and set its value and color.
  // Note that this operation will fail if the active cell is in the top row.
  let aboveCell = activeCell.getOffsetRange(-1, 0);
  aboveCell.setValue("Above cell");
  console.log(`The above cell's address is: ${aboveCell.getAddress()}`);
  aboveCell.getFormat().getFont().setColor("White");
  aboveCell.getFormat().getFill().setColor("Black");
}
```

### <a name="change-all-adjacent-cells"></a><span data-ttu-id="0caa3-122">Изменение всех смежных ячеек</span><span class="sxs-lookup"><span data-stu-id="0caa3-122">Change all adjacent cells</span></span>

<span data-ttu-id="0caa3-123">Этот сценарий копирует форматирование в активной ячейке в соседние ячейки.</span><span class="sxs-lookup"><span data-stu-id="0caa3-123">This script copies the formatting in the active cell to the neighboring cells.</span></span> <span data-ttu-id="0caa3-124">Обратите внимание, что этот скрипт работает только в том случае, если активная ячейка не находится на границе листа.</span><span class="sxs-lookup"><span data-stu-id="0caa3-124">Note that this script only works when the active cell isn't on an edge of the worksheet.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the active cell.
  let activeCell = workbook.getActiveCell();

  // Get the cell that's one row above and one column to the left of the active cell.
  let cornerCell = activeCell.getOffsetRange(-1,-1);

  // Get a range that includes all the cells surrounding the active cell.
  let surroundingRange = cornerCell.getResizedRange(2, 2)

  // Copy the formatting from the active cell to the new range.
  surroundingRange.copyFrom(
    activeCell, /* The source range. */
    ExcelScript.RangeCopyType.formats /* What to copy. */
    );
}
```

### <a name="change-each-individual-cell-in-a-range"></a><span data-ttu-id="0caa3-125">Изменение каждой отдельной ячейки в диапазоне</span><span class="sxs-lookup"><span data-stu-id="0caa3-125">Change each individual cell in a range</span></span>

<span data-ttu-id="0caa3-126">Этот сценарий выполняет цикл над текущим выбранным диапазоном.</span><span class="sxs-lookup"><span data-stu-id="0caa3-126">This script loops over the currently select range.</span></span> <span data-ttu-id="0caa3-127">Он удаляет текущее форматирование и задает случайный цвет для цвета заливки в каждой ячейке.</span><span class="sxs-lookup"><span data-stu-id="0caa3-127">It clears the current formatting and sets the fill color in each cell to a random color.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the currently selected range.
  let range = workbook.getSelectedRange();

  // Get the size boundaries of the range.
  let rows = range.getRowCount();
  let cols = range.getColumnCount();

  // Clear any existing formatting
  range.clear(ExcelScript.ClearApplyTo.formats);

  // Iterate over the range.
  for (let row = 0; row < rows; row++) {
    for (let col = 0; col < cols; col++) {
      // Generate a random color hex-code.
      let colorString = `#${Math.random().toString(16).substr(-6)}`;

      // Set the color of the current cell to that random hex-code.
      range.getCell(row, col).getFormat().getFill().setColor(colorString);
    }
  }
}
```

## <a name="collections"></a><span data-ttu-id="0caa3-128">Коллекции</span><span class="sxs-lookup"><span data-stu-id="0caa3-128">Collections</span></span>

<span data-ttu-id="0caa3-129">Эти примеры работают с коллекциями объектов в книге.</span><span class="sxs-lookup"><span data-stu-id="0caa3-129">These samples work with collections of objects in the workbook.</span></span>

### <a name="iterating-over-collections"></a><span data-ttu-id="0caa3-130">Итерация по коллекциям</span><span class="sxs-lookup"><span data-stu-id="0caa3-130">Iterating over collections</span></span>

<span data-ttu-id="0caa3-131">Этот сценарий получает и заносит в журнал имена всех листов в книге.</span><span class="sxs-lookup"><span data-stu-id="0caa3-131">This script gets and logs the names of all the worksheets in the workbook.</span></span> <span data-ttu-id="0caa3-132">Кроме того, в качестве цвета вкладки задается случайный цвет.</span><span class="sxs-lookup"><span data-stu-id="0caa3-132">It also sets the their tab colors to a random color.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get all the worksheets in the workbook.
  let sheets = workbook.getWorksheets();

  // Get a list of all the worksheet names.
  let names = sheets.map ((sheet) => sheet.getName());

  // Write in the console all the worksheet names and the total count.
  console.log(names);
  console.log(`Total worksheets inside of this workbook: ${sheets.length}`);
  
  // Set the tab color each worksheet to a random color
  for (let sheet of sheets) {
    // Generate a random color hex-code.
    let colorString = `#${Math.random().toString(16).substr(-6)}`;

    // Set the color of the current worksheet's tab to that random hex-code.
    sheet.setTabColor(colorString);
  }
}
```

### <a name="querying-and-deleting-from-a-collection"></a><span data-ttu-id="0caa3-133">Запрос и удаление из коллекции</span><span class="sxs-lookup"><span data-stu-id="0caa3-133">Querying and deleting from a collection</span></span>

<span data-ttu-id="0caa3-134">Этот сценарий создает новый лист.</span><span class="sxs-lookup"><span data-stu-id="0caa3-134">This script creates a new worksheet.</span></span> <span data-ttu-id="0caa3-135">Он проверяет существующую копию листа и удаляет его перед созданием нового листа.</span><span class="sxs-lookup"><span data-stu-id="0caa3-135">It checks for an existing copy of the worksheet and deletes it before making a new sheet.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Name of the worksheet to be added.
  let name = "Index";

  // Get any worksheet with that name.
  let sheet = workbook.getWorksheet("Index");
  
  // If `null` wasn't returned, then there's already a worksheet with that name.
  if (sheet) {
    console.log(`Worksheet by the name ${name} already exists. Deleting it.`);
    // Delete the sheet.
    sheet.delete();
  }
  
  // Add a blank worksheet with the name "Index".
  // Note that this code runs regardless of whether an existing sheet was deleted.
  console.log(`Adding the worksheet named ${name}.`);
  let newSheet = workbook.addWorksheet("Index");

  // Switch to the new worksheet.
  newSheet.activate();
}
```

## <a name="dates"></a><span data-ttu-id="0caa3-136">Даты</span><span class="sxs-lookup"><span data-stu-id="0caa3-136">Dates</span></span>

<span data-ttu-id="0caa3-137">В примерах, приведенных в этом разделе, показано, как использовать объект JavaScript [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) .</span><span class="sxs-lookup"><span data-stu-id="0caa3-137">The samples in this section show how to use the JavaScript [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) object.</span></span>

<span data-ttu-id="0caa3-138">В следующем примере возвращается текущая дата и время, а затем эти значения записываются в две ячейки активного листа.</span><span class="sxs-lookup"><span data-stu-id="0caa3-138">The following sample gets the current date and time and then writes those values to two cells in the active worksheet.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the cells at A1 and B1.
  let dateRange = workbook.getActiveWorksheet().getRange("A1");
  let timeRange = workbook.getActiveWorksheet().getRange("B1");

  // Get the current date and time with the JavaScript Date object.
  let date = new Date(Date.now());

  // Add the date string to A1.
  dateRange.setValue(date.toLocaleDateString());

  // Add the time string to B1.
  timeRange.setValue(date.toLocaleTimeString());
}
```

<span data-ttu-id="0caa3-139">В следующем примере считывается дата, которая хранится в Excel, и преобразуется в объект даты JavaScript.</span><span class="sxs-lookup"><span data-stu-id="0caa3-139">The next sample reads a date that's stored in Excel and translates it to a JavaScript Date object.</span></span> <span data-ttu-id="0caa3-140">В качестве входных данных для даты JavaScript в качестве входных данных используется [числовой серийный номер даты](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) .</span><span class="sxs-lookup"><span data-stu-id="0caa3-140">It uses the [date's numeric serial number](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) as input for the JavaScript Date.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Read a date at cell A1 from Excel.
  let dateRange = workbook.getActiveWorksheet().getRange("A1");

  // Convert the Excel date to a JavaScript Date object.
  let excelDateValue = dateRange.getValue();
  let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
  console.log(javaScriptDate);
}
```

## <a name="display-data"></a><span data-ttu-id="0caa3-141">Отображение данных</span><span class="sxs-lookup"><span data-stu-id="0caa3-141">Display data</span></span>

<span data-ttu-id="0caa3-142">В этих примерах показано, как работать с данными листа и предоставлять пользователям лучшее представление или организацию.</span><span class="sxs-lookup"><span data-stu-id="0caa3-142">These samples demonstrate how to work with worksheet data and provide users with a better view or organization.</span></span>

### <a name="apply-conditional-formatting"></a><span data-ttu-id="0caa3-143">Применение условного форматирования</span><span class="sxs-lookup"><span data-stu-id="0caa3-143">Apply conditional formatting</span></span>

<span data-ttu-id="0caa3-144">В этом примере применяется условное форматирование для диапазона, используемого в текущий момент на листе.</span><span class="sxs-lookup"><span data-stu-id="0caa3-144">This sample applies conditional formatting to the currently used range in the worksheet.</span></span> <span data-ttu-id="0caa3-145">Условное форматирование — Зеленая заливка для первых 10% значений.</span><span class="sxs-lookup"><span data-stu-id="0caa3-145">The conditional formatting is a green fill for the top 10% of values.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the used range in the worksheet.
  let range = selectedSheet.getUsedRange();

  // Set the fill color to green for the top 10% of values in the range.
  let conditionalFormat = range.addConditionalFormat(ExcelScript.ConditionalFormatType.topBottom)
  conditionalFormat.getTopBottom().getFormat().getFill().setColor("green");
  conditionalFormat.getTopBottom().setRule({
    rank: 10, // The percentage threshold.
    type: ExcelScript.ConditionalTopBottomCriterionType.topPercent // The type of the top/bottom condition.
  });
}
```

### <a name="create-a-sorted-table"></a><span data-ttu-id="0caa3-146">Создание отсортированной таблицы</span><span class="sxs-lookup"><span data-stu-id="0caa3-146">Create a sorted table</span></span>

<span data-ttu-id="0caa3-147">В этом примере создается таблица на основе используемого диапазона текущего листа, а затем она сортируется по первому столбцу.</span><span class="sxs-lookup"><span data-stu-id="0caa3-147">This sample creates a table from the current worksheet's used range, then sorts it based on the first column.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Create a table with the used cells.
  let usedRange = selectedSheet.getUsedRange();
  let newTable = selectedSheet.addTable(usedRange, true);

  // Sort the table using the first column.
  newTable.getSort().apply([{ key: 0, ascending: true }]);
}
```

### <a name="log-the-grand-total-values-from-a-pivottable"></a><span data-ttu-id="0caa3-148">Запись значений "общий итог" из сводной таблицы</span><span class="sxs-lookup"><span data-stu-id="0caa3-148">Log the "Grand Total" values from a PivotTable</span></span>

<span data-ttu-id="0caa3-149">В этом примере выполняется поиск первой сводной таблицы в книге и записываются значения в ячейках "общий итог" (выделено зеленым цветом на изображении ниже).</span><span class="sxs-lookup"><span data-stu-id="0caa3-149">This sample finds the first PivotTable in the workbook and logs the values in the "Grand Total" cells (as highlighted in green in the image below).</span></span>

![Сводная таблица продаж фруктов с выделенным зеленым цветом строкой итогов.](../images/sample-pivottable-grand-total-row.png)

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the first PivotTable in the workbook.
  let pivotTable = workbook.getPivotTables()[0];

  // Get the names of each data column in the PivotTable.
  let pivotColumnLabelRange = pivotTable.getLayout().getColumnLabelRange();

  // Get the range displaying the pivoted data.
  let pivotDataRange = pivotTable.getLayout().getRangeBetweenHeaderAndTotal();

  // Get the range with the "grand totals" for the PivotTable columns.
  let grandTotalRange = pivotDataRange.getLastRow();

  // Print each of the "Grand Totals" to the console.
  grandTotalRange.getValues()[0].forEach((column, columnIndex) => {
    console.log(`Grand total of ${pivotColumnLabelRange.getValues()[0][columnIndex]}: ${grandTotalRange.getValues()[0][columnIndex]}`);
    // Example log: "Grand total of Sum of Crates Sold Wholesale: 11000"
  });
}
```

## <a name="formulas"></a><span data-ttu-id="0caa3-151">Формулы</span><span class="sxs-lookup"><span data-stu-id="0caa3-151">Formulas</span></span>

<span data-ttu-id="0caa3-152">В этих примерах используются формулы Excel, а также показано, как работать с ними в скриптах.</span><span class="sxs-lookup"><span data-stu-id="0caa3-152">These samples use Excel formulas and show how to work with them in scripts.</span></span>

## <a name="single-formula"></a><span data-ttu-id="0caa3-153">Одинарная формула</span><span class="sxs-lookup"><span data-stu-id="0caa3-153">Single formula</span></span>

<span data-ttu-id="0caa3-154">Этот сценарий задает формулу ячейки, а затем показывает, как Excel сохраняет формулу и значение ячейки по отдельности.</span><span class="sxs-lookup"><span data-stu-id="0caa3-154">This script sets a cell's formula, then displays how Excel stores the cell's formula and value separately.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
  let selectedSheet = workbook.getActiveWorksheet();

  // Set A1 to 2.
  let a1 = selectedSheet.getRange("A1");
  a1.setValue(2);

  // Set B1 to the formula =(2*A1), which should equal 4.
  let b1 = selectedSheet.getRange("B1")
  b1.setFormula("=(2*A1)");

  // Log the current results for `getFormula` and `getValue` at B1.
  console.log(`B1 - Formula: ${b1.getFormula()} | Value: ${b1.getValue()}`);
}
```

### <a name="spilling-results-from-a-formula"></a><span data-ttu-id="0caa3-155">Сброс результатов из формулы</span><span class="sxs-lookup"><span data-stu-id="0caa3-155">Spilling results from a formula</span></span>

<span data-ttu-id="0caa3-156">Этот сценарий переставит диапазон "a1: D2" на "A4: B7" с помощью функции транспонировать.</span><span class="sxs-lookup"><span data-stu-id="0caa3-156">This script transposes the range "A1:D2" to "A4:B7" by using the TRANSPOSE function.</span></span> <span data-ttu-id="0caa3-157">Если результат передается в #SPILL ошибку, он очищает целевой диапазон и повторно применяет эту формулу.</span><span class="sxs-lookup"><span data-stu-id="0caa3-157">If the transpose results in a #SPILL error, it clears the target range and applies the formula again.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
  let sheet = workbook.getActiveWorksheet();
  // Use the data in A1:D2 for the sample.
  let dataAddress = "A1:D2"
  let inputRange = sheet.getRange(dataAddress);

  // Place the transposed data starting at A4.
  let targetStartCell = sheet.getRange("A4");

  // Compute the target range.
  let targetRange = targetStartCell.getResizedRange(inputRange.getColumnCount() - 1, inputRange.getRowCount() - 1);

  // Call the transpose helper function.
  targetStartCell.setFormula(`=TRANSPOSE(${dataAddress})`);

  // Check if the range update resulted in a spill error.
  let checkValue = targetStartCell.getValue() as string;
  if (checkValue === '#SPILL!') {
    // Clear the target range and call the transpose function again.
    console.log("Target range has data that is preventing update. Clearing target range.");
    targetRange.clear();
    targetStartCell.setFormula(`=TRANSPOSE(${dataAddress})`);
  }

  // Select the transposed range to highlight it.
  targetRange.select();
}
```

## <a name="scenario-samples"></a><span data-ttu-id="0caa3-158">Примеры сценариев</span><span class="sxs-lookup"><span data-stu-id="0caa3-158">Scenario samples</span></span>

<span data-ttu-id="0caa3-159">Примеры, иллюстрирующие большие, реальные решения, можно найти на странице [примеры сценариев для сценариев Office](scenarios/sample-scenario-overview.md).</span><span class="sxs-lookup"><span data-stu-id="0caa3-159">For samples showcasing larger, real-world solutions, visit [Sample scenarios for Office Scripts](scenarios/sample-scenario-overview.md).</span></span>

## <a name="suggest-new-samples"></a><span data-ttu-id="0caa3-160">Предлагаемые новые примеры</span><span class="sxs-lookup"><span data-stu-id="0caa3-160">Suggest new samples</span></span>

<span data-ttu-id="0caa3-161">Мы будем рады получать новые примеры.</span><span class="sxs-lookup"><span data-stu-id="0caa3-161">We welcome suggestions for new samples.</span></span> <span data-ttu-id="0caa3-162">Если существует распространенный сценарий, который поможет другим разработчикам скриптов, Расскажите нам в разделе отзывов, приведенном ниже.</span><span class="sxs-lookup"><span data-stu-id="0caa3-162">If there is a common scenario that would help other script developers, please tell us in the feedback section below.</span></span>
