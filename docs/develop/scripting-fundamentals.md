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
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="f17a4-103">Основные сведения о сценариях для сценариев Office в Excel в Интернете (Предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="f17a4-103">Scripting fundamentals for Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="f17a4-104">В этой статье представлены технические аспекты сценариев Office.</span><span class="sxs-lookup"><span data-stu-id="f17a4-104">This article will introduce you to the technical aspects of Office Scripts.</span></span> <span data-ttu-id="f17a4-105">Вы узнаете, как объекты Excel работают вместе и как редактор кода синхронизируется с книгой.</span><span class="sxs-lookup"><span data-stu-id="f17a4-105">You'll learn how the Excel objects work together and how the Code Editor synchronizes with a workbook.</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="object-model"></a><span data-ttu-id="f17a4-106">Объектная модель</span><span class="sxs-lookup"><span data-stu-id="f17a4-106">Object model</span></span>

<span data-ttu-id="f17a4-107">Чтобы ознакомиться с API Excel, необходимо знать, как компоненты книги связаны друг с другом.</span><span class="sxs-lookup"><span data-stu-id="f17a4-107">To understand the Excel APIs, you must understand how the components of a workbook are related to one another.</span></span>

- <span data-ttu-id="f17a4-108">**Книга** содержит один или несколько **листов**.</span><span class="sxs-lookup"><span data-stu-id="f17a4-108">A **Workbook** contains one or more **Worksheets**.</span></span>
- <span data-ttu-id="f17a4-109">**Лист** предоставляет доступ к ячейкам с помощью объектов **Range** .</span><span class="sxs-lookup"><span data-stu-id="f17a4-109">A **Worksheet** gives access to cells through **Range** objects.</span></span>
- <span data-ttu-id="f17a4-110">**Диапазон** представляет группу смежных ячеек.</span><span class="sxs-lookup"><span data-stu-id="f17a4-110">A **Range** represents a group of contiguous cells.</span></span>
- <span data-ttu-id="f17a4-111">**Диапазоны** используются для создания и размещения **таблиц**, **диаграмм**, **фигур**и других объектов визуализации данных или организации.</span><span class="sxs-lookup"><span data-stu-id="f17a4-111">**Ranges** are used to create and place **Tables**, **Charts**, **Shapes**, and other data visualization or organization objects.</span></span>
- <span data-ttu-id="f17a4-112">**Лист** содержит коллекции объектов данных, присутствующих на отдельном листе.</span><span class="sxs-lookup"><span data-stu-id="f17a4-112">A **Worksheet** contains collections of those data objects that are present in the individual sheet.</span></span>
- <span data-ttu-id="f17a4-113">**Книги** содержат коллекции некоторых объектов данных (например, **таблиц**) для всей **книги**.</span><span class="sxs-lookup"><span data-stu-id="f17a4-113">**Workbooks** contain collections of some of those data objects (such as **Tables**) for the entire **Workbook**.</span></span>

### <a name="ranges"></a><span data-ttu-id="f17a4-114">Ranges</span><span class="sxs-lookup"><span data-stu-id="f17a4-114">Ranges</span></span>

<span data-ttu-id="f17a4-115">Диапазон — это группа смежных ячеек в книге.</span><span class="sxs-lookup"><span data-stu-id="f17a4-115">A range is a group of contiguous cells in the workbook.</span></span> <span data-ttu-id="f17a4-116">В сценариях обычно используется нотация в стиле a1 (например, **B3** для одной ячейки в строке **B** и столбца **3** или **C2: F4** для ячеек из строк **C** и **F** и столбцов **2** – **4**) для определения диапазонов.</span><span class="sxs-lookup"><span data-stu-id="f17a4-116">Scripts typically use A1-style notation (e.g. **B3** for the single cell in row **B** and column **3** or **C2:F4** for the cells from rows **C** through **F** and columns **2** through **4**) to define ranges.</span></span>

<span data-ttu-id="f17a4-117">Диапазоны имеют три основных свойства `values`: `formulas`, и `format`.</span><span class="sxs-lookup"><span data-stu-id="f17a4-117">Ranges have three core properties: `values`, `formulas`, and `format`.</span></span> <span data-ttu-id="f17a4-118">Эти свойства получают или задают значения ячеек, вычисляемые формулы и визуальное форматирование ячеек.</span><span class="sxs-lookup"><span data-stu-id="f17a4-118">These properties get or set the cell values, formulas to be evaluated, and the visual formatting of the cells.</span></span>

#### <a name="range-sample"></a><span data-ttu-id="f17a4-119">Пример диапазона</span><span class="sxs-lookup"><span data-stu-id="f17a4-119">Range sample</span></span>

<span data-ttu-id="f17a4-120">В следующем примере показано, как создать записи о продажах.</span><span class="sxs-lookup"><span data-stu-id="f17a4-120">The following sample shows how to create sales records.</span></span> <span data-ttu-id="f17a4-121">В этом сценарии `Range` используются объекты для задания значений, формул и форматов.</span><span class="sxs-lookup"><span data-stu-id="f17a4-121">This script uses `Range` objects to set the values, formulas, and formats.</span></span>

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

<span data-ttu-id="f17a4-122">При выполнении этого сценария на текущем листе создаются следующие данные:</span><span class="sxs-lookup"><span data-stu-id="f17a4-122">Running this script creates the following data in the current worksheet:</span></span>

![Запись о продажах, в которой показаны строки значений, столбец формул и отформатированные заголовки.](../images/range-sample.png)

### <a name="charts-tables-and-other-data-objects"></a><span data-ttu-id="f17a4-124">Диаграммы, таблицы и другие объекты данных</span><span class="sxs-lookup"><span data-stu-id="f17a4-124">Charts, tables, and other data objects</span></span>

<span data-ttu-id="f17a4-125">Скрипты могут создавать и управлять структурами данных и зрительными представлениями в Excel.</span><span class="sxs-lookup"><span data-stu-id="f17a4-125">Scripts can create and manipulate the data structures and visualizations within Excel.</span></span> <span data-ttu-id="f17a4-126">Таблицы и диаграммы — это два наиболее часто используемых объекта, но API поддерживают сводные таблицы, фигуры, изображения и многое другое.</span><span class="sxs-lookup"><span data-stu-id="f17a4-126">Tables and charts are two of the more commonly used objects, but the APIs support PivotTables, shapes, images, and more.</span></span>

#### <a name="creating-a-table"></a><span data-ttu-id="f17a4-127">Создание таблицы</span><span class="sxs-lookup"><span data-stu-id="f17a4-127">Creating a table</span></span>

<span data-ttu-id="f17a4-128">Создание таблиц с использованием диапазонов, заполненных данными.</span><span class="sxs-lookup"><span data-stu-id="f17a4-128">Create tables by using data-filled ranges.</span></span> <span data-ttu-id="f17a4-129">Элементы управления форматированием и таблицами (например, фильтры) автоматически применяются к диапазону.</span><span class="sxs-lookup"><span data-stu-id="f17a4-129">Formatting and table controls (such as filters) are automatically applied to the range.</span></span>

<span data-ttu-id="f17a4-130">Следующий сценарий создает таблицу, используя диапазоны из предыдущего примера.</span><span class="sxs-lookup"><span data-stu-id="f17a4-130">The following script creates a table using the ranges from the previous sample.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
   let sheet = context.workbook.worksheets.getActiveWorksheet();
   sheet.tables.add("B2:E5", true);
}
```

<span data-ttu-id="f17a4-131">При выполнении этого скрипта на листе с предыдущими данными создается следующая таблица:</span><span class="sxs-lookup"><span data-stu-id="f17a4-131">Running this script on the worksheet with the previous data creates the following table:</span></span>

![Таблица, созданная из предыдущей записи о продажах.](../images/table-sample.png)

#### <a name="creating-a-chart"></a><span data-ttu-id="f17a4-133">Создание диаграммы</span><span class="sxs-lookup"><span data-stu-id="f17a4-133">Creating a chart</span></span>

<span data-ttu-id="f17a4-134">Создание диаграмм для отображения данных в диапазоне.</span><span class="sxs-lookup"><span data-stu-id="f17a4-134">Create charts to visualize the data in a range.</span></span> <span data-ttu-id="f17a4-135">Сценарии позволяют использовать десятки видов диаграмм, каждая из которых может быть изменена в соответствии со своими потребностями.</span><span class="sxs-lookup"><span data-stu-id="f17a4-135">Scripts allow for dozens of chart varieties, each of which can be customized to suit your needs.</span></span>

<span data-ttu-id="f17a4-136">Следующий сценарий создает простую гистограмму для трех элементов и размещает его на 100 пикселов ниже верхней границы листа.</span><span class="sxs-lookup"><span data-stu-id="f17a4-136">The following script creates a simple column chart for three items and places it 100 pixels below the top of the worksheet.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
  let sheet = context.workbook.worksheets.getActiveWorksheet();
  let chart = sheet.charts.add(Excel.ChartType.columnStacked, sheet.getRange("B3:C5"));
  chart.top = 100;
}
```

<span data-ttu-id="f17a4-137">При выполнении этого сценария на листе с помощью предыдущей таблицы создается следующая диаграмма:</span><span class="sxs-lookup"><span data-stu-id="f17a4-137">Running this script on the worksheet with the previous table creates the following chart:</span></span>

![Гистограмма, на которой показано количество трех элементов из предыдущей записи о продажах.](../images/chart-sample.png)

### <a name="further-reading-on-the-object-model"></a><span data-ttu-id="f17a4-139">Дальнейшие материалы по объектной модели</span><span class="sxs-lookup"><span data-stu-id="f17a4-139">Further reading on the object model</span></span>

<span data-ttu-id="f17a4-140">[Справочная документация по API сценариев Office](/javascript/api/office-scripts/overview) — это полный список объектов, используемых в сценариях Office.</span><span class="sxs-lookup"><span data-stu-id="f17a4-140">The [Office Scripts API reference documentation](/javascript/api/office-scripts/overview) is a comprehensive listing of the objects used in Office Scripts.</span></span> <span data-ttu-id="f17a4-141">Здесь вы можете использовать оглавление, чтобы перейти к любому классу, о котором вы хотите узнать больше.</span><span class="sxs-lookup"><span data-stu-id="f17a4-141">There, you can use the table of contents to navigate to any class you'd like to learn more about.</span></span> <span data-ttu-id="f17a4-142">Ниже представлено несколько часто просматриваемых страниц.</span><span class="sxs-lookup"><span data-stu-id="f17a4-142">The following are several commonly viewed pages.</span></span>

- [<span data-ttu-id="f17a4-143">Chart</span><span class="sxs-lookup"><span data-stu-id="f17a4-143">Chart</span></span>](/javascript/api/office-scripts/excel/excel.chart)
- [<span data-ttu-id="f17a4-144">Comment</span><span class="sxs-lookup"><span data-stu-id="f17a4-144">Comment</span></span>](/javascript/api/office-scripts/excel/excel.comment)
- [<span data-ttu-id="f17a4-145">PivotTable</span><span class="sxs-lookup"><span data-stu-id="f17a4-145">PivotTable</span></span>](/javascript/api/office-scripts/excel/excel.pivottable)
- [<span data-ttu-id="f17a4-146">Range</span><span class="sxs-lookup"><span data-stu-id="f17a4-146">Range</span></span>](/javascript/api/office-scripts/excel/excel.range)
- [<span data-ttu-id="f17a4-147">RangeFormat</span><span class="sxs-lookup"><span data-stu-id="f17a4-147">RangeFormat</span></span>](/javascript/api/office-scripts/excel/excel.rangeformat)
- [<span data-ttu-id="f17a4-148">Shape</span><span class="sxs-lookup"><span data-stu-id="f17a4-148">Shape</span></span>](/javascript/api/office-scripts/excel/excel.shape)
- [<span data-ttu-id="f17a4-149">Table</span><span class="sxs-lookup"><span data-stu-id="f17a4-149">Table</span></span>](/javascript/api/office-scripts/excel/excel.table)
- [<span data-ttu-id="f17a4-150">Workbook</span><span class="sxs-lookup"><span data-stu-id="f17a4-150">Workbook</span></span>](/javascript/api/office-scripts/excel/excel.workbook)
- [<span data-ttu-id="f17a4-151">Worksheet</span><span class="sxs-lookup"><span data-stu-id="f17a4-151">Worksheet</span></span>](/javascript/api/office-scripts/excel/excel.worksheet)

## <a name="main-function"></a><span data-ttu-id="f17a4-152">`main`функциями</span><span class="sxs-lookup"><span data-stu-id="f17a4-152">`main` function</span></span>

<span data-ttu-id="f17a4-153">Каждый сценарий Office должен содержать `main` функцию со следующей подписью, в том числе `Excel.RequestContext` определение типа:</span><span class="sxs-lookup"><span data-stu-id="f17a4-153">Every Office Script must contain a `main` function with the following signature, including the `Excel.RequestContext` type definition:</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your Excel Script
}
```

<span data-ttu-id="f17a4-154">Код внутри `main` функции выполняется при запуске скрипта.</span><span class="sxs-lookup"><span data-stu-id="f17a4-154">The code inside the `main` function runs when the script is run.</span></span> <span data-ttu-id="f17a4-155">`main`может вызывать другие функции в вашем сценарии, но код, не содержащийся в функции, не будет выполняться.</span><span class="sxs-lookup"><span data-stu-id="f17a4-155">`main` can call other functions in your script, but code that's not contained in a function will not run.</span></span>

## <a name="context"></a><span data-ttu-id="f17a4-156">Контекст</span><span class="sxs-lookup"><span data-stu-id="f17a4-156">Context</span></span>

<span data-ttu-id="f17a4-157">`main` Функция принимает `Excel.RequestContext` параметр с именем `context`.</span><span class="sxs-lookup"><span data-stu-id="f17a4-157">The `main` function accepts an `Excel.RequestContext` parameter, named `context`.</span></span> <span data-ttu-id="f17a4-158">Подумайте о `context` качестве моста между сценарием и книгой.</span><span class="sxs-lookup"><span data-stu-id="f17a4-158">Think of `context` as the bridge between your script and the workbook.</span></span> <span data-ttu-id="f17a4-159">Ваш скрипт получает доступ к книге с `context` объектом и использует его `context` для отправки данных назад и вперед.</span><span class="sxs-lookup"><span data-stu-id="f17a4-159">Your script accesses the workbook with the `context` object and uses that `context` to send data back and forth.</span></span>

<span data-ttu-id="f17a4-160">Этот `context` объект необходим, так как скрипт и Excel выполняются в различных процессах и расположениях.</span><span class="sxs-lookup"><span data-stu-id="f17a4-160">The `context` object is necessary because the script and Excel are running in different processes and locations.</span></span> <span data-ttu-id="f17a4-161">Скрипту потребуется внести изменения в данные или запросить данные из книги в облаке.</span><span class="sxs-lookup"><span data-stu-id="f17a4-161">The script will need to make changes to or query data from the workbook in the cloud.</span></span> <span data-ttu-id="f17a4-162">`context` Объект управляет этими транзакциями.</span><span class="sxs-lookup"><span data-stu-id="f17a4-162">The `context` object manages those transactions.</span></span>

## <a name="sync-and-load"></a><span data-ttu-id="f17a4-163">Синхронизация и загрузка</span><span class="sxs-lookup"><span data-stu-id="f17a4-163">Sync and Load</span></span>

<span data-ttu-id="f17a4-164">Так как сценарий и книга работают в различных расположениях, передача данных между ними занимает некоторое время.</span><span class="sxs-lookup"><span data-stu-id="f17a4-164">Because your script and workbook run in different locations, any data transfer between the two takes time.</span></span> <span data-ttu-id="f17a4-165">Для повышения производительности скриптов команды ставятся в очередь до тех пор, пока скрипт не `sync` выполнит явное вызов операции для синхронизации скрипта и книги.</span><span class="sxs-lookup"><span data-stu-id="f17a4-165">To improve script performance, commands are queued up until the script explicitly calls the `sync` operation to synchronize the script and workbook.</span></span> <span data-ttu-id="f17a4-166">Ваш сценарий может работать независимо, пока ему не потребуется выполнить одно из следующих действий:</span><span class="sxs-lookup"><span data-stu-id="f17a4-166">Your script can work independently until it needs to do either of the following:</span></span>

- <span data-ttu-id="f17a4-167">Чтение данных из книги (после `load` операции).</span><span class="sxs-lookup"><span data-stu-id="f17a4-167">Read data from the workbook (following a `load` operation).</span></span>
- <span data-ttu-id="f17a4-168">Запись данных в книгу (обычно потому, что сценарий завершен).</span><span class="sxs-lookup"><span data-stu-id="f17a4-168">Write data to the workbook (usually because the script has finished).</span></span>

<span data-ttu-id="f17a4-169">На приведенном ниже рисунке показан пример последовательности элементов управления между скриптом и книгой:</span><span class="sxs-lookup"><span data-stu-id="f17a4-169">The following image shows an example control flow between the script and workbook:</span></span>

![Схема, демонстрирующая операции чтения и записи, идущие к книге из скрипта.](../images/load-sync.png)

### <a name="sync"></a><span data-ttu-id="f17a4-171">Синхронизировать</span><span class="sxs-lookup"><span data-stu-id="f17a4-171">Sync</span></span>

<span data-ttu-id="f17a4-172">Если сценарий должен считывать данные из книги или записывать данные в нее, вызовите `RequestContext.sync` метод, как показано ниже:</span><span class="sxs-lookup"><span data-stu-id="f17a4-172">Whenever your script needs to read data from or write data to the workbook, call the `RequestContext.sync` method as shown here:</span></span>

```TypeScript
await context.sync();
```

> [!NOTE]
> <span data-ttu-id="f17a4-173">`context.sync()`вызывается неявно при завершении скрипта.</span><span class="sxs-lookup"><span data-stu-id="f17a4-173">`context.sync()` is implicitly called when a script ends.</span></span>

<span data-ttu-id="f17a4-174">После завершения `sync` операции книга обновляется, чтобы отразить все операции записи, заданные в скрипте.</span><span class="sxs-lookup"><span data-stu-id="f17a4-174">After the `sync` operation completes, the workbook updates to reflect any write operations that script has specified.</span></span> <span data-ttu-id="f17a4-175">Операция записи задает любое свойство объекта Excel (например, `range.format.fill.color = "red"`) или вызывает метод, который изменяет свойство (например, `range.format.autoFitColumns()`).</span><span class="sxs-lookup"><span data-stu-id="f17a4-175">A write operation is setting any property on a Excel object (e.g. `range.format.fill.color = "red"`) or calling a method that changes a property (e.g., `range.format.autoFitColumns()`).</span></span> <span data-ttu-id="f17a4-176">Эта `sync` операция также считывает все значения из книги, запрошенные сценарием с помощью `load` операции (как описано в следующем разделе).</span><span class="sxs-lookup"><span data-stu-id="f17a4-176">The `sync` operation also reads any values from the workbook that the script requested by using a `load` operation (as discussed in the next section).</span></span>

<span data-ttu-id="f17a4-177">Синхронизация вашего сценария с книгой может занять некоторое время в зависимости от сети.</span><span class="sxs-lookup"><span data-stu-id="f17a4-177">Synchronizing your script with the workbook can take time, depending on your network.</span></span> <span data-ttu-id="f17a4-178">Необходимо минимизировать количество `sync` вызовов, чтобы ускорить выполнение сценария.</span><span class="sxs-lookup"><span data-stu-id="f17a4-178">You should minimize the number of `sync` calls to help your script run fast.</span></span>  

### <a name="load"></a><span data-ttu-id="f17a4-179">Load</span><span class="sxs-lookup"><span data-stu-id="f17a4-179">Load</span></span>

<span data-ttu-id="f17a4-180">Перед прочтением скрипту необходимо загрузить данные из книги.</span><span class="sxs-lookup"><span data-stu-id="f17a4-180">A script must load data from the workbook before reading it.</span></span> <span data-ttu-id="f17a4-181">Однако часто загрузка данных из всей книги значительно сокращает скорость сценария.</span><span class="sxs-lookup"><span data-stu-id="f17a4-181">However, frequently loading data from the entire workbook would greatly reduce the script's speed.</span></span> <span data-ttu-id="f17a4-182">Вместо этого `load` метод предоставляет состояние скрипта, в частности, какие данные следует извлечь из книги.</span><span class="sxs-lookup"><span data-stu-id="f17a4-182">Instead, the `load` method lets your script state specifically which data should be retrieved from the workbook.</span></span>

<span data-ttu-id="f17a4-183">`load` Метод доступен для каждого объекта Excel.</span><span class="sxs-lookup"><span data-stu-id="f17a4-183">The `load` method is available on every Excel object.</span></span> <span data-ttu-id="f17a4-184">Скрипту необходимо загрузить свойства объекта, прежде чем он сможет прочитать их.</span><span class="sxs-lookup"><span data-stu-id="f17a4-184">Your script must load an object's properties before it can read them.</span></span> <span data-ttu-id="f17a4-185">В противном случае будет возникать ошибка.</span><span class="sxs-lookup"><span data-stu-id="f17a4-185">Not doing so will result in an error.</span></span>

<span data-ttu-id="f17a4-186">В приведенных ниже примерах используется `Range` объект для отображения трех способов `load` , которые метод может использовать для загрузки данных.</span><span class="sxs-lookup"><span data-stu-id="f17a4-186">The following examples use a `Range` object to show the three ways the `load` method can be used to load data.</span></span>

|<span data-ttu-id="f17a4-187">Условие</span><span class="sxs-lookup"><span data-stu-id="f17a4-187">Intent</span></span> |<span data-ttu-id="f17a4-188">Пример команды</span><span class="sxs-lookup"><span data-stu-id="f17a4-188">Example Command</span></span> | <span data-ttu-id="f17a4-189">Эффект</span><span class="sxs-lookup"><span data-stu-id="f17a4-189">Effect</span></span> |
|:--|:--|:--|
|<span data-ttu-id="f17a4-190">Загрузить одно свойство</span><span class="sxs-lookup"><span data-stu-id="f17a4-190">Load one property</span></span> |`myRange.load("values");` | <span data-ttu-id="f17a4-191">Загружает одно свойство, в данном случае двухмерный массив значений из этого диапазона.</span><span class="sxs-lookup"><span data-stu-id="f17a4-191">Loads a single property, in this case the two-dimensional array of values in this range.</span></span> |
|<span data-ttu-id="f17a4-192">Загрузка нескольких свойств</span><span class="sxs-lookup"><span data-stu-id="f17a4-192">Load multiple properties</span></span> |`myRange.load("values, rowCount, columnCount");`| <span data-ttu-id="f17a4-193">Загружает все свойства из списка, разделенного запятыми, в данном примере значения, количество строк и число столбцов.</span><span class="sxs-lookup"><span data-stu-id="f17a4-193">Loads all the properties from a comma-delimited list, in this example the values, row count, and column count.</span></span> |
|<span data-ttu-id="f17a4-194">Загрузка всех элементов</span><span class="sxs-lookup"><span data-stu-id="f17a4-194">Load everything</span></span> | `myRange.load();`|<span data-ttu-id="f17a4-195">Загружает все свойства диапазона.</span><span class="sxs-lookup"><span data-stu-id="f17a4-195">Loads all the properties on the range.</span></span> <span data-ttu-id="f17a4-196">Это не рекомендуемое решение, так как оно замедляет выполнение скрипта, получая ненужные данные.</span><span class="sxs-lookup"><span data-stu-id="f17a4-196">This is not a recommended solution, since it will slow down your script by getting unnecessary data.</span></span> <span data-ttu-id="f17a4-197">Этот параметр следует использовать только при тестировании скрипта или при необходимости каждого свойства объекта.</span><span class="sxs-lookup"><span data-stu-id="f17a4-197">You should only use this while testing your script or if you need every property from the object.</span></span> |

<span data-ttu-id="f17a4-198">Перед чтением любых загруженных значений ваш сценарий должен вызываться `context.sync()` .</span><span class="sxs-lookup"><span data-stu-id="f17a4-198">Your script must call `context.sync()` before reading any loaded values.</span></span>

```TypeScript
let range = selectedSheet.getRange("A1:B3");
range.load ("rowCount"); // Load the property.
await context.sync(); // Synchronize with the workbook to get the property.
console.log(range.rowCount); // Read and log the property value (3).
```

<span data-ttu-id="f17a4-199">Кроме того, можно загружать свойства во всей коллекции.</span><span class="sxs-lookup"><span data-stu-id="f17a4-199">You can also load properties across an entire collection.</span></span> <span data-ttu-id="f17a4-200">У каждого объекта Collection есть `items` свойство, которое представляет собой массив, содержащий объекты в этой коллекции.</span><span class="sxs-lookup"><span data-stu-id="f17a4-200">Every collection object has an `items` property that is an array containing the objects in that collection.</span></span> <span data-ttu-id="f17a4-201">Использование `items` в качестве начала иерархического вызова (`items\myProperty`) для `load` загрузки указанных свойств для каждого из этих элементов.</span><span class="sxs-lookup"><span data-stu-id="f17a4-201">Using `items` as the start of a hierarchical call (`items\myProperty`) to `load` loads the specified properties on each of those items.</span></span> <span data-ttu-id="f17a4-202">В следующем примере загружается `resolved` свойство для `Comment` каждого объекта в `CommentCollection` объекте листа.</span><span class="sxs-lookup"><span data-stu-id="f17a4-202">The following example loads the `resolved` property on every `Comment` object in the `CommentCollection` object of a worksheet.</span></span>

```TypeScript
let comments = selectedSheet.comments;
comments.load("items/resolved"); // Load the `resolved` property from every comment in this collection.
await context.sync(); // Synchronize with the workbook to get the properties.
```

> [!TIP]
> <span data-ttu-id="f17a4-203">Чтобы узнать больше о работе с коллекциями в сценариях Office, ознакомьтесь с [разделом Array, посвященным встроенным объектам JavaScript в статье сценариев Office](javascript-objects.md#array) .</span><span class="sxs-lookup"><span data-stu-id="f17a4-203">To learn more about working with collections in Office Scripts, see the [Array section of the Using built-in JavaScript objects in Office Scripts](javascript-objects.md#array) article.</span></span>

## <a name="see-also"></a><span data-ttu-id="f17a4-204">См. также</span><span class="sxs-lookup"><span data-stu-id="f17a4-204">See also</span></span>

- [<span data-ttu-id="f17a4-205">Запись, редактирование и создание сценариев Office в Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="f17a4-205">Record, edit, and create Office Scripts in Excel on the web</span></span>](../tutorials/excel-tutorial.md)
- [<span data-ttu-id="f17a4-206">Чтение данных книги с помощью скриптов Office в Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="f17a4-206">Read workbook data with Office Scripts in Excel on the web</span></span>](../tutorials/excel-read-tutorial.md)
- [<span data-ttu-id="f17a4-207">Справочник по API скриптов Office</span><span class="sxs-lookup"><span data-stu-id="f17a4-207">Office Scripts API reference</span></span>](/javascript/api/office-scripts/overview)
- [<span data-ttu-id="f17a4-208">Использование встроенных объектов JavaScript в сценариях Office</span><span class="sxs-lookup"><span data-stu-id="f17a4-208">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)
