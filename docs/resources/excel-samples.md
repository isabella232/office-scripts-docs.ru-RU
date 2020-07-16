---
title: Примеры сценариев для сценариев Office в Excel в Интернете
description: Коллекция примеров кода для использования со сценариями Office в Excel в Интернете.
ms.date: 06/18/2020
localization_priority: Normal
ms.openlocfilehash: bfa6679595e6e28cc5d2ae3e3e487fd3e77738aa
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878677"
---
# <a name="sample-scripts-for-office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="9ff59-103">Примеры сценариев для сценариев Office в Excel в Интернете (Предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="9ff59-103">Sample scripts for Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="9ff59-104">Ниже приведены примеры простых сценариев, которые можно использовать в собственных книгах.</span><span class="sxs-lookup"><span data-stu-id="9ff59-104">The following samples are simple scripts for you to try on your own workbooks.</span></span> <span data-ttu-id="9ff59-105">Чтобы использовать их в Excel в Интернете, выполните следующие действия:</span><span class="sxs-lookup"><span data-stu-id="9ff59-105">To use them in Excel on the web:</span></span>

1. <span data-ttu-id="9ff59-106">Откройте вкладку **Автоматизировать**.</span><span class="sxs-lookup"><span data-stu-id="9ff59-106">Open the **Automate** tab.</span></span>
2. <span data-ttu-id="9ff59-107">Нажмите клавишу **Редактор кода**.</span><span class="sxs-lookup"><span data-stu-id="9ff59-107">Press **Code Editor**.</span></span>
3. <span data-ttu-id="9ff59-108">Нажмите **новый скрипт** в области задач редактора кода.</span><span class="sxs-lookup"><span data-stu-id="9ff59-108">Press **New Script** in the Code Editor's task pane.</span></span>
4. <span data-ttu-id="9ff59-109">Замените весь сценарий выбранным образцом.</span><span class="sxs-lookup"><span data-stu-id="9ff59-109">Replace the entire script with the sample of your choice.</span></span>
5. <span data-ttu-id="9ff59-110">В области задач редактора кода нажмите кнопку **запустить** .</span><span class="sxs-lookup"><span data-stu-id="9ff59-110">Press **Run** in the Code Editor's task pane.</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="scripting-basics"></a><span data-ttu-id="9ff59-111">Основные сведения о сценариях</span><span class="sxs-lookup"><span data-stu-id="9ff59-111">Scripting basics</span></span>

<span data-ttu-id="9ff59-112">В этих примерах демонстрируются основные конструктивные блоки для сценариев Office.</span><span class="sxs-lookup"><span data-stu-id="9ff59-112">These samples demonstrate fundamental building blocks for Office Scripts.</span></span> <span data-ttu-id="9ff59-113">Добавьте их в скрипты, чтобы расширить решение и устранить распространенные проблемы.</span><span class="sxs-lookup"><span data-stu-id="9ff59-113">Add these to your scripts to extend your solution and solve common problems.</span></span>

### <a name="read-and-log-one-cell"></a><span data-ttu-id="9ff59-114">Чтение и запись в журнал одной ячейки</span><span class="sxs-lookup"><span data-stu-id="9ff59-114">Read and log one cell</span></span>

<span data-ttu-id="9ff59-115">В этом примере считывается значение **a1** и выводится на консоль.</span><span class="sxs-lookup"><span data-stu-id="9ff59-115">This sample reads the value of **A1** and prints it to the console.</span></span>

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

### <a name="read-the-active-cell"></a><span data-ttu-id="9ff59-116">Чтение активной ячейки</span><span class="sxs-lookup"><span data-stu-id="9ff59-116">Read the active cell</span></span>

<span data-ttu-id="9ff59-117">Этот сценарий записывает в журнал значение текущей активной ячейки.</span><span class="sxs-lookup"><span data-stu-id="9ff59-117">This script logs the value of the current active cell.</span></span> <span data-ttu-id="9ff59-118">Если выбрано несколько ячеек, в журнал заносится левая верхняя ячейка.</span><span class="sxs-lookup"><span data-stu-id="9ff59-118">If multiple cells are selected, the top-leftmost cell will be logged.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the current active cell in the workbook.
  let cell = workbook.getActiveCell();

  // Log that cell's value.
  console.log(`The current cell's value is ${cell.getValue()}`);
}
```

### <a name="change-an-adjacent-cell"></a><span data-ttu-id="9ff59-119">Изменение смежной ячейки</span><span class="sxs-lookup"><span data-stu-id="9ff59-119">Change an adjacent cell</span></span>

<span data-ttu-id="9ff59-120">Этот сценарий получает смежные ячейки, используя относительные ссылки.</span><span class="sxs-lookup"><span data-stu-id="9ff59-120">This script gets adjacent cells using relative references.</span></span> <span data-ttu-id="9ff59-121">Обратите внимание, что если активная ячейка находится в верхней строке, часть скрипта завершается с ошибкой, так как она ссылается на ячейку над выбранной в текущий момент.</span><span class="sxs-lookup"><span data-stu-id="9ff59-121">Note that if the active cell is on the top row, part of the script fails, because it references the cell above the currently selected one.</span></span>

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

### <a name="change-all-adjacent-cells"></a><span data-ttu-id="9ff59-122">Изменение всех смежных ячеек</span><span class="sxs-lookup"><span data-stu-id="9ff59-122">Change all adjacent cells</span></span>

<span data-ttu-id="9ff59-123">Этот сценарий копирует форматирование в активной ячейке в соседние ячейки.</span><span class="sxs-lookup"><span data-stu-id="9ff59-123">This script copies the formatting in the active cell to the neighboring cells.</span></span> <span data-ttu-id="9ff59-124">Обратите внимание, что этот скрипт работает только в том случае, если активная ячейка не находится на границе листа.</span><span class="sxs-lookup"><span data-stu-id="9ff59-124">Note that this script only works when the active cell isn't on an edge of the worksheet.</span></span>

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

### <a name="work-with-dates"></a><span data-ttu-id="9ff59-125">Работать с датами</span><span class="sxs-lookup"><span data-stu-id="9ff59-125">Work with dates</span></span>

<span data-ttu-id="9ff59-126">В примерах, приведенных в этом разделе, показано, как использовать объект JavaScript [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) .</span><span class="sxs-lookup"><span data-stu-id="9ff59-126">The samples in this section show how to use the JavaScript [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) object.</span></span>

<span data-ttu-id="9ff59-127">В следующем примере возвращается текущая дата и время, а затем эти значения записываются в две ячейки активного листа.</span><span class="sxs-lookup"><span data-stu-id="9ff59-127">The following sample gets the current date and time and then writes those values to two cells in the active worksheet.</span></span>

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

<span data-ttu-id="9ff59-128">В следующем примере считывается дата, которая хранится в Excel, и преобразуется в объект даты JavaScript.</span><span class="sxs-lookup"><span data-stu-id="9ff59-128">The next sample reads a date that's stored in Excel and translates it to a JavaScript Date object.</span></span> <span data-ttu-id="9ff59-129">В качестве входных данных для даты JavaScript в качестве входных данных используется [числовой серийный номер даты](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) .</span><span class="sxs-lookup"><span data-stu-id="9ff59-129">It uses the [date's numeric serial number](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) as input for the JavaScript Date.</span></span>

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

## <a name="display-data"></a><span data-ttu-id="9ff59-130">Отображение данных</span><span class="sxs-lookup"><span data-stu-id="9ff59-130">Display data</span></span>

<span data-ttu-id="9ff59-131">В этих примерах показано, как работать с данными листа и предоставлять пользователям лучшее представление или организацию.</span><span class="sxs-lookup"><span data-stu-id="9ff59-131">These samples demonstrate how to work with worksheet data and provide users with a better view or organization.</span></span>

### <a name="apply-conditional-formatting"></a><span data-ttu-id="9ff59-132">Применение условного форматирования</span><span class="sxs-lookup"><span data-stu-id="9ff59-132">Apply conditional formatting</span></span>

<span data-ttu-id="9ff59-133">В этом примере применяется условное форматирование для диапазона, используемого в текущий момент на листе.</span><span class="sxs-lookup"><span data-stu-id="9ff59-133">This sample applies conditional formatting to the currently used range in the worksheet.</span></span> <span data-ttu-id="9ff59-134">Условное форматирование — Зеленая заливка для первых 10% значений.</span><span class="sxs-lookup"><span data-stu-id="9ff59-134">The conditional formatting is a green fill for the top 10% of values.</span></span>

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

### <a name="create-a-sorted-table"></a><span data-ttu-id="9ff59-135">Создание отсортированной таблицы</span><span class="sxs-lookup"><span data-stu-id="9ff59-135">Create a sorted table</span></span>

<span data-ttu-id="9ff59-136">В этом примере создается таблица на основе используемого диапазона текущего листа, а затем она сортируется по первому столбцу.</span><span class="sxs-lookup"><span data-stu-id="9ff59-136">This sample creates a table from the current worksheet's used range, then sorts it based on the first column.</span></span>

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

### <a name="log-the-grand-total-values-from-a-pivottable"></a><span data-ttu-id="9ff59-137">Запись значений "общий итог" из сводной таблицы</span><span class="sxs-lookup"><span data-stu-id="9ff59-137">Log the "Grand Total" values from a PivotTable</span></span>

<span data-ttu-id="9ff59-138">В этом примере выполняется поиск первой сводной таблицы в книге и записываются значения в ячейках "общий итог" (выделено зеленым цветом на изображении ниже).</span><span class="sxs-lookup"><span data-stu-id="9ff59-138">This sample finds the first PivotTable in the workbook and logs the values in the "Grand Total" cells (as highlighted in green in the image below).</span></span>

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

## <a name="scenario-samples"></a><span data-ttu-id="9ff59-140">Примеры сценариев</span><span class="sxs-lookup"><span data-stu-id="9ff59-140">Scenario samples</span></span>

<span data-ttu-id="9ff59-141">Примеры, иллюстрирующие большие, реальные решения, можно найти на странице [примеры сценариев для сценариев Office](scenarios/sample-scenario-overview.md).</span><span class="sxs-lookup"><span data-stu-id="9ff59-141">For samples showcasing larger, real-world solutions, visit [Sample scenarios for Office Scripts](scenarios/sample-scenario-overview.md).</span></span>

## <a name="suggest-new-samples"></a><span data-ttu-id="9ff59-142">Предлагаемые новые примеры</span><span class="sxs-lookup"><span data-stu-id="9ff59-142">Suggest new samples</span></span>

<span data-ttu-id="9ff59-143">Мы будем рады получать новые примеры.</span><span class="sxs-lookup"><span data-stu-id="9ff59-143">We welcome suggestions for new samples.</span></span> <span data-ttu-id="9ff59-144">Если существует распространенный сценарий, который поможет другим разработчикам скриптов, Расскажите нам в разделе отзывов, приведенном ниже.</span><span class="sxs-lookup"><span data-stu-id="9ff59-144">If there is a common scenario that would help other script developers, please tell us in the feedback section below.</span></span>
