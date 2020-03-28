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
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="c5f62-103">Основы сценариев для сценариев Office в Excel в Интернете (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="c5f62-103">Scripting fundamentals for Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="c5f62-104">Эта статья познакомит вас с техническими аспектами сценариев Office.</span><span class="sxs-lookup"><span data-stu-id="c5f62-104">This article will introduce you to the technical aspects of Office Scripts.</span></span> <span data-ttu-id="c5f62-105">Вы узнаете, как объекты Excel работают вместе и как редактор кода синхронизируется с книгой.</span><span class="sxs-lookup"><span data-stu-id="c5f62-105">You'll learn how the Excel objects work together and how the Code Editor synchronizes with a workbook.</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="object-model"></a><span data-ttu-id="c5f62-106">Объектная модель</span><span class="sxs-lookup"><span data-stu-id="c5f62-106">Object model</span></span>

<span data-ttu-id="c5f62-107">Чтобы понять API-интерфейсы Excel, вы должны понимать, как компоненты рабочей книги связаны друг с другом.</span><span class="sxs-lookup"><span data-stu-id="c5f62-107">To understand the Excel APIs, you must understand how the components of a workbook are related to one another.</span></span>

- <span data-ttu-id="c5f62-108">**Рабочая книга** содержит одну или несколько **рабочих листов**.</span><span class="sxs-lookup"><span data-stu-id="c5f62-108">A **Workbook** contains one or more **Worksheets**.</span></span>
- <span data-ttu-id="c5f62-109">**Рабочий лист** предоставляет доступ к ячейкам через объекты **Range**.</span><span class="sxs-lookup"><span data-stu-id="c5f62-109">A **Worksheet** gives access to cells through **Range** objects.</span></span>
- <span data-ttu-id="c5f62-110">**Range** представляет группу смежных клеток.</span><span class="sxs-lookup"><span data-stu-id="c5f62-110">A **Range** represents a group of contiguous cells.</span></span>
- <span data-ttu-id="c5f62-111">**Диапазоны** используются для создания и размещения **таблиц**, **диаграмм**, **фигур** и других объектов визуализации данных или организации.</span><span class="sxs-lookup"><span data-stu-id="c5f62-111">**Ranges** are used to create and place **Tables**, **Charts**, **Shapes**, and other data visualization or organization objects.</span></span>
- <span data-ttu-id="c5f62-112">**Рабочий лист** содержит коллекции тех объектов данных, которые присутствуют на отдельном листе.</span><span class="sxs-lookup"><span data-stu-id="c5f62-112">A **Worksheet** contains collections of those data objects that are present in the individual sheet.</span></span>
- <span data-ttu-id="c5f62-113">**Рабочие книги** содержат коллекции некоторых из этих объектов данных (таких как **таблицы**) для всей **рабочей книги**.</span><span class="sxs-lookup"><span data-stu-id="c5f62-113">**Workbooks** contain collections of some of those data objects (such as **Tables**) for the entire **Workbook**.</span></span>

### <a name="ranges"></a><span data-ttu-id="c5f62-114">Диапазоны</span><span class="sxs-lookup"><span data-stu-id="c5f62-114">Ranges</span></span>

<span data-ttu-id="c5f62-115">Диапазон - это группа непрерывных ячеек в рабочей книге.</span><span class="sxs-lookup"><span data-stu-id="c5f62-115">A range is a group of contiguous cells in the workbook.</span></span> <span data-ttu-id="c5f62-116">В сценариях обычно используется нотация в стиле A1 (например, **B3** для отдельной ячейки в строке **B** и столбце **3** или **C2:F4** для ячеек из строк с **C** по **F** и столбцов со **2** по **4**) для определения диапазонов.</span><span class="sxs-lookup"><span data-stu-id="c5f62-116">Scripts typically use A1-style notation (e.g. **B3** for the single cell in row **B** and column **3** or **C2:F4** for the cells from rows **C** through **F** and columns **2** through **4**) to define ranges.</span></span>

<span data-ttu-id="c5f62-117">Диапазоны имеют три основных свойства: `values`, `formulas`, и `format`.</span><span class="sxs-lookup"><span data-stu-id="c5f62-117">Ranges have three core properties: `values`, `formulas`, and `format`.</span></span> <span data-ttu-id="c5f62-118">Эти свойства получают или устанавливают значения ячеек, формулы для оценки и визуальное форматирование ячеек.</span><span class="sxs-lookup"><span data-stu-id="c5f62-118">These properties get or set the cell values, formulas to be evaluated, and the visual formatting of the cells.</span></span>

#### <a name="range-sample"></a><span data-ttu-id="c5f62-119">Образец диапазона</span><span class="sxs-lookup"><span data-stu-id="c5f62-119">Range sample</span></span>

<span data-ttu-id="c5f62-120">В следующем примере показано, как создавать записи продаж.</span><span class="sxs-lookup"><span data-stu-id="c5f62-120">The following sample shows how to create sales records.</span></span> <span data-ttu-id="c5f62-121">Этот скрипт использует `Range` объекты для установки значений, формул и форматов.</span><span class="sxs-lookup"><span data-stu-id="c5f62-121">This script uses `Range` objects to set the values, formulas, and formats.</span></span>

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

<span data-ttu-id="c5f62-122">Выполнение этого скрипта создает следующие данные в текущей рабочей таблице:</span><span class="sxs-lookup"><span data-stu-id="c5f62-122">Running this script creates the following data in the current worksheet:</span></span>

![Запись о продажах, показывающая строки значений, столбец формулы и отформатированные заголовки.](../images/range-sample.png)

### <a name="charts-tables-and-other-data-objects"></a><span data-ttu-id="c5f62-124">Диаграммы, таблицы и другие объекты данных</span><span class="sxs-lookup"><span data-stu-id="c5f62-124">Charts, tables, and other data objects</span></span>

<span data-ttu-id="c5f62-125">Скрипты могут создавать и управлять структурами данных и визуализациями в Excel.</span><span class="sxs-lookup"><span data-stu-id="c5f62-125">Scripts can create and manipulate the data structures and visualizations within Excel.</span></span> <span data-ttu-id="c5f62-126">Таблицы и диаграммы являются двумя наиболее часто используемыми объектами, но API поддерживают сводные таблицы, фигуры, изображения и многое другое.</span><span class="sxs-lookup"><span data-stu-id="c5f62-126">Tables and charts are two of the more commonly used objects, but the APIs support PivotTables, shapes, images, and more.</span></span>

#### <a name="creating-a-table"></a><span data-ttu-id="c5f62-127">Создание таблицы</span><span class="sxs-lookup"><span data-stu-id="c5f62-127">Creating a table</span></span>

<span data-ttu-id="c5f62-128">Создавайте таблицы, используя заполненные данными диапазоны.</span><span class="sxs-lookup"><span data-stu-id="c5f62-128">Create tables by using data-filled ranges.</span></span> <span data-ttu-id="c5f62-129">Элементы управления форматированием и таблицами (например, фильтры) автоматически применяются к диапазону.</span><span class="sxs-lookup"><span data-stu-id="c5f62-129">Formatting and table controls (such as filters) are automatically applied to the range.</span></span>

<span data-ttu-id="c5f62-130">Следующий скрипт создает таблицу с использованием диапазонов из предыдущего примера.</span><span class="sxs-lookup"><span data-stu-id="c5f62-130">The following script creates a table using the ranges from the previous sample.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
   let sheet = context.workbook.worksheets.getActiveWorksheet();
   sheet.tables.add("B2:E5", true);
}
```

<span data-ttu-id="c5f62-131">Выполнение этого сценария на листе с предыдущими данными создает следующую таблицу:</span><span class="sxs-lookup"><span data-stu-id="c5f62-131">Running this script on the worksheet with the previous data creates the following table:</span></span>

![Таблица сделана из предыдущего рекорда продаж.](../images/table-sample.png)

#### <a name="creating-a-chart"></a><span data-ttu-id="c5f62-133">Создание диаграммы</span><span class="sxs-lookup"><span data-stu-id="c5f62-133">Creating a chart</span></span>

<span data-ttu-id="c5f62-134">Создайте диаграммы для визуализации данных в диапазоне.</span><span class="sxs-lookup"><span data-stu-id="c5f62-134">Create charts to visualize the data in a range.</span></span> <span data-ttu-id="c5f62-135">Сценарии позволяют создавать десятки разновидностей диаграмм, каждая из которых может быть настроена в соответствии с вашими потребностями.</span><span class="sxs-lookup"><span data-stu-id="c5f62-135">Scripts allow for dozens of chart varieties, each of which can be customized to suit your needs.</span></span>

<span data-ttu-id="c5f62-136">Следующий скрипт создает простую столбчатую диаграмму для трех элементов и размещает ее на 100 пикселей ниже верхней части листа.</span><span class="sxs-lookup"><span data-stu-id="c5f62-136">The following script creates a simple column chart for three items and places it 100 pixels below the top of the worksheet.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
  let sheet = context.workbook.worksheets.getActiveWorksheet();
  let chart = sheet.charts.add(Excel.ChartType.columnStacked, sheet.getRange("B3:C5"));
  chart.top = 100;
}
```

<span data-ttu-id="c5f62-137">Запуск этого скрипта на листе с предыдущей таблицей создает следующую диаграмму:</span><span class="sxs-lookup"><span data-stu-id="c5f62-137">Running this script on the worksheet with the previous table creates the following chart:</span></span>

![Столбчатая диаграмма, показывающая количество трех предметов из предыдущей записи о продажах.](../images/chart-sample.png)

### <a name="further-reading-on-the-object-model"></a><span data-ttu-id="c5f62-139">Дальнейшее чтение по объектной модели</span><span class="sxs-lookup"><span data-stu-id="c5f62-139">Further reading on the object model</span></span>

<span data-ttu-id="c5f62-140">[Справочная документация по API сценариев Office](/javascript/api/office-scripts/overview) представляет собой полный список объектов, используемых в сценариях Office.</span><span class="sxs-lookup"><span data-stu-id="c5f62-140">The [Office Scripts API reference documentation](/javascript/api/office-scripts/overview) is a comprehensive listing of the objects used in Office Scripts.</span></span> <span data-ttu-id="c5f62-141">Там вы можете использовать оглавление, чтобы перейти к любому классу, о котором вы хотите узнать больше.</span><span class="sxs-lookup"><span data-stu-id="c5f62-141">There, you can use the table of contents to navigate to any class you'd like to learn more about.</span></span> <span data-ttu-id="c5f62-142">Ниже приведены несколько часто просматриваемых страниц.</span><span class="sxs-lookup"><span data-stu-id="c5f62-142">The following are several commonly viewed pages.</span></span>

- [<span data-ttu-id="c5f62-143">Chart</span><span class="sxs-lookup"><span data-stu-id="c5f62-143">Chart</span></span>](/javascript/api/office-scripts/excel/excel.chart)
- [<span data-ttu-id="c5f62-144">Comment</span><span class="sxs-lookup"><span data-stu-id="c5f62-144">Comment</span></span>](/javascript/api/office-scripts/excel/excel.comment)
- [<span data-ttu-id="c5f62-145">PivotTable</span><span class="sxs-lookup"><span data-stu-id="c5f62-145">PivotTable</span></span>](/javascript/api/office-scripts/excel/excel.pivottable)
- [<span data-ttu-id="c5f62-146">Range</span><span class="sxs-lookup"><span data-stu-id="c5f62-146">Range</span></span>](/javascript/api/office-scripts/excel/excel.range)
- [<span data-ttu-id="c5f62-147">RangeFormat</span><span class="sxs-lookup"><span data-stu-id="c5f62-147">RangeFormat</span></span>](/javascript/api/office-scripts/excel/excel.rangeformat)
- [<span data-ttu-id="c5f62-148">Shape</span><span class="sxs-lookup"><span data-stu-id="c5f62-148">Shape</span></span>](/javascript/api/office-scripts/excel/excel.shape)
- [<span data-ttu-id="c5f62-149">Table</span><span class="sxs-lookup"><span data-stu-id="c5f62-149">Table</span></span>](/javascript/api/office-scripts/excel/excel.table)
- [<span data-ttu-id="c5f62-150">Workbook</span><span class="sxs-lookup"><span data-stu-id="c5f62-150">Workbook</span></span>](/javascript/api/office-scripts/excel/excel.workbook)
- [<span data-ttu-id="c5f62-151">Worksheet</span><span class="sxs-lookup"><span data-stu-id="c5f62-151">Worksheet</span></span>](/javascript/api/office-scripts/excel/excel.worksheet)

## <a name="main-function"></a><span data-ttu-id="c5f62-152">`main` функция</span><span class="sxs-lookup"><span data-stu-id="c5f62-152">`main` function</span></span>

<span data-ttu-id="c5f62-153">Каждый сценарий Office должен содержать `main` функцию со следующей подписью, включая определение `Excel.RequestContext` типа:</span><span class="sxs-lookup"><span data-stu-id="c5f62-153">Every Office Script must contain a `main` function with the following signature, including the `Excel.RequestContext` type definition:</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your Excel Script
}
```

<span data-ttu-id="c5f62-154">Код внутри `main` функции запускается при запуске скрипта.</span><span class="sxs-lookup"><span data-stu-id="c5f62-154">The code inside the `main` function runs when the script is run.</span></span> <span data-ttu-id="c5f62-155">`main` может вызывать другие функции в вашем скрипте, но код, который не содержится в функции, не будет работать.</span><span class="sxs-lookup"><span data-stu-id="c5f62-155">`main` can call other functions in your script, but code that's not contained in a function will not run.</span></span>

## <a name="context"></a><span data-ttu-id="c5f62-156">Context</span><span class="sxs-lookup"><span data-stu-id="c5f62-156">Context</span></span>

<span data-ttu-id="c5f62-157">Функция `main` принимает `Excel.RequestContext` параметра с именем `context`.</span><span class="sxs-lookup"><span data-stu-id="c5f62-157">The `main` function accepts an `Excel.RequestContext` parameter, named `context`.</span></span> <span data-ttu-id="c5f62-158">Думайте о `context` как о мосте между вашим сценарием и книгой.</span><span class="sxs-lookup"><span data-stu-id="c5f62-158">Think of `context` as the bridge between your script and the workbook.</span></span> <span data-ttu-id="c5f62-159">Ваш сценарий обращается к книге с помощью `context` объекта и использует этот `context` для отправки данных туда и обратно.</span><span class="sxs-lookup"><span data-stu-id="c5f62-159">Your script accesses the workbook with the `context` object and uses that `context` to send data back and forth.</span></span>

<span data-ttu-id="c5f62-160">Объект `context` необходим, потому что скрипт и Excel работают в разных процессах и местах.</span><span class="sxs-lookup"><span data-stu-id="c5f62-160">The `context` object is necessary because the script and Excel are running in different processes and locations.</span></span> <span data-ttu-id="c5f62-161">Сценарий должен будет внести изменения или запросить данные из рабочей книги в облаке.</span><span class="sxs-lookup"><span data-stu-id="c5f62-161">The script will need to make changes to or query data from the workbook in the cloud.</span></span> <span data-ttu-id="c5f62-162">Объект `context` управляет этими транзакциями.</span><span class="sxs-lookup"><span data-stu-id="c5f62-162">The `context` object manages those transactions.</span></span>

## <a name="sync-and-load"></a><span data-ttu-id="c5f62-163">Синхронизация и загрузка</span><span class="sxs-lookup"><span data-stu-id="c5f62-163">Sync and Load</span></span>

<span data-ttu-id="c5f62-164">Поскольку ваш сценарий и рабочая книга работают в разных местах, любая передача данных между ними занимает много времени.</span><span class="sxs-lookup"><span data-stu-id="c5f62-164">Because your script and workbook run in different locations, any data transfer between the two takes time.</span></span> <span data-ttu-id="c5f62-165">Для повышения производительности сценария команды помещаются в очередь до тех пор, пока сценарий явно не вызовет `sync` операцию для синхронизации сценария и рабочей книги.</span><span class="sxs-lookup"><span data-stu-id="c5f62-165">To improve script performance, commands are queued up until the script explicitly calls the `sync` operation to synchronize the script and workbook.</span></span> <span data-ttu-id="c5f62-166">Ваш скрипт может работать независимо, пока он не выполнит одно из следующих действий:</span><span class="sxs-lookup"><span data-stu-id="c5f62-166">Your script can work independently until it needs to do either of the following:</span></span>

- <span data-ttu-id="c5f62-167">Чтение данных из рабочей книги (после `load` операции).</span><span class="sxs-lookup"><span data-stu-id="c5f62-167">Read data from the workbook (following a `load` operation).</span></span>
- <span data-ttu-id="c5f62-168">Запишите данные в рабочую книгу (обычно потому, что сценарий завершен).</span><span class="sxs-lookup"><span data-stu-id="c5f62-168">Write data to the workbook (usually because the script has finished).</span></span>

<span data-ttu-id="c5f62-169">На следующем рисунке показан пример потока управления между сценарием и книгой:</span><span class="sxs-lookup"><span data-stu-id="c5f62-169">The following image shows an example control flow between the script and workbook:</span></span>

![Диаграмма, показывающая операции чтения и записи, идущие в рабочую книгу из сценария.](../images/load-sync.png)

### <a name="sync"></a><span data-ttu-id="c5f62-171">Синхронизировать</span><span class="sxs-lookup"><span data-stu-id="c5f62-171">Sync</span></span>

<span data-ttu-id="c5f62-172">Всякий раз, когда вашему сценарию нужно прочитать данные или записать данные в рабочую книгу, вызывайте метод `RequestContext.sync`, как показано здесь:</span><span class="sxs-lookup"><span data-stu-id="c5f62-172">Whenever your script needs to read data from or write data to the workbook, call the `RequestContext.sync` method as shown here:</span></span>

```TypeScript
await context.sync();
```

> [!NOTE]
> <span data-ttu-id="c5f62-173">`context.sync()` неявно вызывается, когда скрипт заканчивается.</span><span class="sxs-lookup"><span data-stu-id="c5f62-173">`context.sync()` is implicitly called when a script ends.</span></span>

<span data-ttu-id="c5f62-174">После завершения операции `sync` книга обновляется, чтобы отразить все операции записи, указанные сценарием.</span><span class="sxs-lookup"><span data-stu-id="c5f62-174">After the `sync` operation completes, the workbook updates to reflect any write operations that script has specified.</span></span> <span data-ttu-id="c5f62-175">Операция записи устанавливает любое свойство для объекта Excel (например, `range.format.fill.color = "red"`) или вызывает метод, который изменяет свойство (например, `range.format.autoFitColumns()`).</span><span class="sxs-lookup"><span data-stu-id="c5f62-175">A write operation is setting any property on a Excel object (e.g. `range.format.fill.color = "red"`) or calling a method that changes a property (e.g., `range.format.autoFitColumns()`).</span></span> <span data-ttu-id="c5f62-176">Операция `sync` также считывает любые значения из рабочей книги, запрошенные сценарием с помощью операции `load` (как описано в следующем разделе).</span><span class="sxs-lookup"><span data-stu-id="c5f62-176">The `sync` operation also reads any values from the workbook that the script requested by using a `load` operation (as discussed in the next section).</span></span>

<span data-ttu-id="c5f62-177">Синхронизация вашего сценария с книгой может занять некоторое время, в зависимости от вашей сети.</span><span class="sxs-lookup"><span data-stu-id="c5f62-177">Synchronizing your script with the workbook can take time, depending on your network.</span></span> <span data-ttu-id="c5f62-178">Вы должны минимизировать количество вызовов `sync`, чтобы ваш скрипт работал быстро.</span><span class="sxs-lookup"><span data-stu-id="c5f62-178">You should minimize the number of `sync` calls to help your script run fast.</span></span>  

### <a name="load"></a><span data-ttu-id="c5f62-179">Load</span><span class="sxs-lookup"><span data-stu-id="c5f62-179">Load</span></span>

<span data-ttu-id="c5f62-180">Сценарий должен загрузить данные из рабочей книги перед ее чтением.</span><span class="sxs-lookup"><span data-stu-id="c5f62-180">A script must load data from the workbook before reading it.</span></span> <span data-ttu-id="c5f62-181">Однако частая загрузка данных из всей рабочей книги значительно снижает скорость работы сценария.</span><span class="sxs-lookup"><span data-stu-id="c5f62-181">However, frequently loading data from the entire workbook would greatly reduce the script's speed.</span></span> <span data-ttu-id="c5f62-182">Вместо этого метод `load` позволяет вашему сценарию указать, какие именно данные следует извлечь из рабочей книги.</span><span class="sxs-lookup"><span data-stu-id="c5f62-182">Instead, the `load` method lets your script state specifically which data should be retrieved from the workbook.</span></span>

<span data-ttu-id="c5f62-183">Метод `load` доступен для каждого объекта Excel.</span><span class="sxs-lookup"><span data-stu-id="c5f62-183">The `load` method is available on every Excel object.</span></span> <span data-ttu-id="c5f62-184">Ваш скрипт должен загрузить свойства объекта, прежде чем он сможет их прочитать.</span><span class="sxs-lookup"><span data-stu-id="c5f62-184">Your script must load an object's properties before it can read them.</span></span> <span data-ttu-id="c5f62-185">Невыполнение этого требования приведет к ошибке.</span><span class="sxs-lookup"><span data-stu-id="c5f62-185">Not doing so will result in an error.</span></span>

<span data-ttu-id="c5f62-186">В следующих примерах объект `Range` используется для демонстрации трех способов использования метода `load` для загрузки данных.</span><span class="sxs-lookup"><span data-stu-id="c5f62-186">The following examples use a `Range` object to show the three ways the `load` method can be used to load data.</span></span>

|<span data-ttu-id="c5f62-187">Intent</span><span class="sxs-lookup"><span data-stu-id="c5f62-187">Intent</span></span> |<span data-ttu-id="c5f62-188">Пример команды</span><span class="sxs-lookup"><span data-stu-id="c5f62-188">Example Command</span></span> | <span data-ttu-id="c5f62-189">Эффект</span><span class="sxs-lookup"><span data-stu-id="c5f62-189">Effect</span></span> |
|:--|:--|:--|
|<span data-ttu-id="c5f62-190">Загрузить одно свойство</span><span class="sxs-lookup"><span data-stu-id="c5f62-190">Load one property</span></span> |`myRange.load("values");` | <span data-ttu-id="c5f62-191">Загружает одно свойство, в данном случае двумерный массив значений в этом диапазоне.</span><span class="sxs-lookup"><span data-stu-id="c5f62-191">Loads a single property, in this case the two-dimensional array of values in this range.</span></span> |
|<span data-ttu-id="c5f62-192">Загрузить несколько свойств</span><span class="sxs-lookup"><span data-stu-id="c5f62-192">Load multiple properties</span></span> |`myRange.load("values, rowCount, columnCount");`| <span data-ttu-id="c5f62-193">Загружает все свойства из списка, разделенного запятыми, в этом примере значения, количество строк и количество столбцов.</span><span class="sxs-lookup"><span data-stu-id="c5f62-193">Loads all the properties from a comma-delimited list, in this example the values, row count, and column count.</span></span> |
|<span data-ttu-id="c5f62-194">Загрузить все</span><span class="sxs-lookup"><span data-stu-id="c5f62-194">Load everything</span></span> | `myRange.load();`|<span data-ttu-id="c5f62-195">Загружает все свойства в диапазоне.</span><span class="sxs-lookup"><span data-stu-id="c5f62-195">Loads all the properties on the range.</span></span> <span data-ttu-id="c5f62-196">Это не рекомендуемое решение, так как оно замедлит ваш скрипт, получая ненужные данные.</span><span class="sxs-lookup"><span data-stu-id="c5f62-196">This is not a recommended solution, since it will slow down your script by getting unnecessary data.</span></span> <span data-ttu-id="c5f62-197">Вы должны использовать это только при тестировании вашего скрипта или если вам нужно каждое свойство объекта.</span><span class="sxs-lookup"><span data-stu-id="c5f62-197">You should only use this while testing your script or if you need every property from the object.</span></span> |

<span data-ttu-id="c5f62-198">Ваш скрипт должен вызывать `context.sync()` перед чтением любых загруженных значений.</span><span class="sxs-lookup"><span data-stu-id="c5f62-198">Your script must call `context.sync()` before reading any loaded values.</span></span>

```TypeScript
let range = selectedSheet.getRange("A1:B3");
range.load ("rowCount"); // Load the property.
await context.sync(); // Synchronize with the workbook to get the property.
console.log(range.rowCount); // Read and log the property value (3).
```

<span data-ttu-id="c5f62-199">Вы также можете загрузить свойства всей коллекции.</span><span class="sxs-lookup"><span data-stu-id="c5f62-199">You can also load properties across an entire collection.</span></span> <span data-ttu-id="c5f62-200">Каждый объект коллекции имеет `items` свойство, которое является массивом, содержащим объекты в этой коллекции.</span><span class="sxs-lookup"><span data-stu-id="c5f62-200">Every collection object has an `items` property that is an array containing the objects in that collection.</span></span> <span data-ttu-id="c5f62-201">Использование `items` в качестве начала иерархического вызова (`items\myProperty`) для `load` загружает указанные свойства для каждого из этих элементов.</span><span class="sxs-lookup"><span data-stu-id="c5f62-201">Using `items` as the start of a hierarchical call (`items\myProperty`) to `load` loads the specified properties on each of those items.</span></span> <span data-ttu-id="c5f62-202">В следующем примере загружается свойство `resolved` для каждых `Comment` объектов в `CommentCollection` объекте рабочего листа.</span><span class="sxs-lookup"><span data-stu-id="c5f62-202">The following example loads the `resolved` property on every `Comment` object in the `CommentCollection` object of a worksheet.</span></span>

```TypeScript
let comments = selectedSheet.comments;
comments.load("items/resolved"); // Load the `resolved` property from every comment in this collection.
await context.sync(); // Synchronize with the workbook to get the properties.
```

> [!TIP]
> <span data-ttu-id="c5f62-203">Подробнее о работе с коллекциями в сценариях Office см. в статье, в разделе [массива Использование встроенных объектов JavaScript в сценариях Office](javascript-objects.md#array).</span><span class="sxs-lookup"><span data-stu-id="c5f62-203">To learn more about working with collections in Office Scripts, see the [Array section of the Using built-in JavaScript objects in Office Scripts](javascript-objects.md#array) article.</span></span>

## <a name="see-also"></a><span data-ttu-id="c5f62-204">См. также</span><span class="sxs-lookup"><span data-stu-id="c5f62-204">See also</span></span>

- [<span data-ttu-id="c5f62-205">Запись, редактирование и создание сценариев Office в Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="c5f62-205">Record, edit, and create Office Scripts in Excel on the web</span></span>](../tutorials/excel-tutorial.md)
- [<span data-ttu-id="c5f62-206">Чтение данных рабочей книги с помощью сценариев Office в Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="c5f62-206">Read workbook data with Office Scripts in Excel on the web</span></span>](../tutorials/excel-read-tutorial.md)
- [<span data-ttu-id="c5f62-207">Справочник API для сценариев Office</span><span class="sxs-lookup"><span data-stu-id="c5f62-207">Office Scripts API reference</span></span>](/javascript/api/office-scripts/overview)
- [<span data-ttu-id="c5f62-208">Использование встроенных объектов JavaScript в сценариях Office</span><span class="sxs-lookup"><span data-stu-id="c5f62-208">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)
