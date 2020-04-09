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
# <a name="sample-scripts-for-office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="aa2fb-103">Примеры сценариев для сценариев Office в Excel в Интернете (Предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="aa2fb-103">Sample scripts for Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="aa2fb-104">Ниже приведены примеры простых сценариев, которые можно использовать в собственных книгах.</span><span class="sxs-lookup"><span data-stu-id="aa2fb-104">The following samples are simple scripts for you to try on your own workbooks.</span></span> <span data-ttu-id="aa2fb-105">Чтобы использовать их в Excel в Интернете, выполните следующие действия:</span><span class="sxs-lookup"><span data-stu-id="aa2fb-105">To use them in Excel on the web:</span></span>

1. <span data-ttu-id="aa2fb-106">Откройте вкладку **Автоматизировать**.</span><span class="sxs-lookup"><span data-stu-id="aa2fb-106">Open the **Automate** tab.</span></span>
2. <span data-ttu-id="aa2fb-107">Нажмите клавишу **Редактор кода**.</span><span class="sxs-lookup"><span data-stu-id="aa2fb-107">Press **Code Editor**.</span></span>
3. <span data-ttu-id="aa2fb-108">Нажмите **новый скрипт** в области задач редактора кода.</span><span class="sxs-lookup"><span data-stu-id="aa2fb-108">Press **New Script** in the Code Editor's task pane.</span></span>
4. <span data-ttu-id="aa2fb-109">Замените весь сценарий выбранным образцом.</span><span class="sxs-lookup"><span data-stu-id="aa2fb-109">Replace the entire script with the sample of your choice.</span></span>
5. <span data-ttu-id="aa2fb-110">В области задач редактора кода нажмите кнопку **запустить** .</span><span class="sxs-lookup"><span data-stu-id="aa2fb-110">Press **Run** in the Code Editor's task pane.</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="scripting-basics"></a><span data-ttu-id="aa2fb-111">Основные сведения о сценариях</span><span class="sxs-lookup"><span data-stu-id="aa2fb-111">Scripting basics</span></span>

<span data-ttu-id="aa2fb-112">В этих примерах демонстрируются основные конструктивные блоки для сценариев Office.</span><span class="sxs-lookup"><span data-stu-id="aa2fb-112">These samples demonstrate fundamental building blocks for Office Scripts.</span></span> <span data-ttu-id="aa2fb-113">Добавьте их в скрипты, чтобы расширить решение и устранить распространенные проблемы.</span><span class="sxs-lookup"><span data-stu-id="aa2fb-113">Add these to your scripts to extend your solution and solve common problems.</span></span>

### <a name="read-and-log-one-cell"></a><span data-ttu-id="aa2fb-114">Чтение и запись в журнал одной ячейки</span><span class="sxs-lookup"><span data-stu-id="aa2fb-114">Read and log one cell</span></span>

<span data-ttu-id="aa2fb-115">В этом примере считывается значение **a1** и выводится на консоль.</span><span class="sxs-lookup"><span data-stu-id="aa2fb-115">This sample reads the value of **A1** and prints it to the console.</span></span>

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

### <a name="work-with-dates"></a><span data-ttu-id="aa2fb-116">Работать с датами</span><span class="sxs-lookup"><span data-stu-id="aa2fb-116">Work with dates</span></span>

<span data-ttu-id="aa2fb-117">В примерах, приведенных в этом разделе, показано, как использовать объект JavaScript [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) .</span><span class="sxs-lookup"><span data-stu-id="aa2fb-117">The samples in this section show how to use the JavaScript [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) object.</span></span>

<span data-ttu-id="aa2fb-118">В следующем примере возвращается текущая дата и время, а затем эти значения записываются в две ячейки активного листа.</span><span class="sxs-lookup"><span data-stu-id="aa2fb-118">The following sample gets the current date and time and then writes those values to two cells in the active worksheet.</span></span>

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

<span data-ttu-id="aa2fb-119">В следующем примере считывается дата, которая хранится в Excel, и преобразуется в объект даты JavaScript.</span><span class="sxs-lookup"><span data-stu-id="aa2fb-119">The next sample reads a date that's stored in Excel and translates it to a JavaScript Date object.</span></span> <span data-ttu-id="aa2fb-120">В качестве входных данных для даты JavaScript в качестве входных данных используется [числовой серийный номер даты](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) .</span><span class="sxs-lookup"><span data-stu-id="aa2fb-120">It uses the [date's numeric serial number](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) as input for the JavaScript Date.</span></span>

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

## <a name="display-data"></a><span data-ttu-id="aa2fb-121">Отображение данных</span><span class="sxs-lookup"><span data-stu-id="aa2fb-121">Display data</span></span>

<span data-ttu-id="aa2fb-122">В этих примерах показано, как работать с данными листа и предоставлять пользователям лучшее представление или организацию.</span><span class="sxs-lookup"><span data-stu-id="aa2fb-122">These samples demonstrate how to work with worksheet data and provide users with a better view or organization.</span></span>

### <a name="apply-conditional-formatting"></a><span data-ttu-id="aa2fb-123">Применение условного форматирования</span><span class="sxs-lookup"><span data-stu-id="aa2fb-123">Apply conditional formatting</span></span>

<span data-ttu-id="aa2fb-124">В этом примере применяется условное форматирование для диапазона, используемого в текущий момент на листе.</span><span class="sxs-lookup"><span data-stu-id="aa2fb-124">This sample applies conditional formatting to the currently used range in the worksheet.</span></span> <span data-ttu-id="aa2fb-125">Условное форматирование — Зеленая заливка для первых 10% значений.</span><span class="sxs-lookup"><span data-stu-id="aa2fb-125">The conditional formatting is a green fill for the top 10% of values.</span></span>

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

### <a name="create-a-sorted-table"></a><span data-ttu-id="aa2fb-126">Создание отсортированной таблицы</span><span class="sxs-lookup"><span data-stu-id="aa2fb-126">Create a sorted table</span></span>

<span data-ttu-id="aa2fb-127">В этом примере создается таблица на основе используемого диапазона текущего листа, а затем она сортируется по первому столбцу.</span><span class="sxs-lookup"><span data-stu-id="aa2fb-127">This sample creates a table from the current worksheet's used range, then sorts it based on the first column.</span></span>

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

## <a name="collaboration"></a><span data-ttu-id="aa2fb-128">Совместная работа</span><span class="sxs-lookup"><span data-stu-id="aa2fb-128">Collaboration</span></span>

<span data-ttu-id="aa2fb-129">В этих примерах показано, как работать с функциями Excel, относящимися к совместной работе, например комментариями.</span><span class="sxs-lookup"><span data-stu-id="aa2fb-129">These samples demonstrate how to work with collaboration-related features of Excel, such as comments.</span></span>

### <a name="delete-resolved-comments"></a><span data-ttu-id="aa2fb-130">Удаление разрешенных комментариев</span><span class="sxs-lookup"><span data-stu-id="aa2fb-130">Delete resolved comments</span></span>

<span data-ttu-id="aa2fb-131">В этом примере удаляются все разрешенные комментарии из текущего листа.</span><span class="sxs-lookup"><span data-stu-id="aa2fb-131">This sample deletes all resolved comments from the current worksheet.</span></span>

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

## <a name="scenario-samples"></a><span data-ttu-id="aa2fb-132">Примеры сценариев</span><span class="sxs-lookup"><span data-stu-id="aa2fb-132">Scenario samples</span></span>

<span data-ttu-id="aa2fb-133">Примеры, иллюстрирующие большие, реальные решения, можно найти на странице [примеры сценариев для сценариев Office](scenarios/sample-scenario-overview.md).</span><span class="sxs-lookup"><span data-stu-id="aa2fb-133">For samples showcasing larger, real-world solutions, visit [Sample scenarios for Office Scripts](scenarios/sample-scenario-overview.md).</span></span>

## <a name="suggest-new-samples"></a><span data-ttu-id="aa2fb-134">Предлагаемые новые примеры</span><span class="sxs-lookup"><span data-stu-id="aa2fb-134">Suggest new samples</span></span>

<span data-ttu-id="aa2fb-135">Мы будем рады получать новые примеры.</span><span class="sxs-lookup"><span data-stu-id="aa2fb-135">We welcome suggestions for new samples.</span></span> <span data-ttu-id="aa2fb-136">Если существует распространенный сценарий, который поможет другим разработчикам скриптов, Расскажите нам в разделе отзывов, приведенном ниже.</span><span class="sxs-lookup"><span data-stu-id="aa2fb-136">If there is a common scenario that would help other script developers, please tell us in the feedback section below.</span></span>
