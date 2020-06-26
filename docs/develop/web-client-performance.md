---
title: Повышение производительности сценариев Office
description: Создавайте более быстрые сценарии, чтобы понять связь между книгой Excel и сценарием.
ms.date: 06/15/2020
localization_priority: Normal
ms.openlocfilehash: 4d5b7c70f14e3fc598b95a6226e3ef8caf89f651
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878900"
---
# <a name="improve-the-performance-of-your-office-scripts"></a><span data-ttu-id="3eb98-103">Повышение производительности сценариев Office</span><span class="sxs-lookup"><span data-stu-id="3eb98-103">Improve the performance of your Office Scripts</span></span>

<span data-ttu-id="3eb98-104">Сценарии Office предназначены для автоматизации часто выполняемого ряда задач для экономии времени.</span><span class="sxs-lookup"><span data-stu-id="3eb98-104">The purpose of Office Scripts is to automate commonly performed series of tasks to save you time.</span></span> <span data-ttu-id="3eb98-105">Медленный сценарий может быть очень похожим, так как он не ускоряет рабочий процесс.</span><span class="sxs-lookup"><span data-stu-id="3eb98-105">A slow script can feel like it doesn't speed up your workflow.</span></span> <span data-ttu-id="3eb98-106">В большинстве случаев ваш сценарий прекрасно подходит и выполняется должным образом.</span><span class="sxs-lookup"><span data-stu-id="3eb98-106">Most of the time, your script will be perfectly fine and run as expected.</span></span> <span data-ttu-id="3eb98-107">Однако существует несколько проблем, которые могут повлиять на производительность.</span><span class="sxs-lookup"><span data-stu-id="3eb98-107">However, there are a few, avoidable scenarios that can affect performance.</span></span>

<span data-ttu-id="3eb98-108">Наиболее распространенная причина медленного сценария заключается в чрезмерном обмене данными с книгой.</span><span class="sxs-lookup"><span data-stu-id="3eb98-108">The most common reason for a slow script is excessive communication with the workbook.</span></span> <span data-ttu-id="3eb98-109">Ваш сценарий выполняется на локальном компьютере, а книга существует в облаке.</span><span class="sxs-lookup"><span data-stu-id="3eb98-109">Your script runs on your local machine, while the workbook exists in the cloud.</span></span> <span data-ttu-id="3eb98-110">В определенное время сценарий синхронизирует свои локальные данные с книгой.</span><span class="sxs-lookup"><span data-stu-id="3eb98-110">At certain times, your script synchronizes its local data with that of the workbook.</span></span> <span data-ttu-id="3eb98-111">Это означает, что любые операции записи (например, `workbook.addWorksheet()` ) применяются к книге только при выполнении этой фоновой синхронизации.</span><span class="sxs-lookup"><span data-stu-id="3eb98-111">This means that any write operations (such as `workbook.addWorksheet()`) are only applied to the workbook when this behind-the-scenes synchronization happens.</span></span> <span data-ttu-id="3eb98-112">Аналогично, любые операции чтения (например, `myRange.getValues()` ) получают данные только из книги для сценария в указанное время.</span><span class="sxs-lookup"><span data-stu-id="3eb98-112">Likewise, any read operations (such as `myRange.getValues()`) only get data from the workbook for the script at those times.</span></span> <span data-ttu-id="3eb98-113">В обоих случаях скрипт получает сведения, прежде чем они действуют с данными.</span><span class="sxs-lookup"><span data-stu-id="3eb98-113">In either case, the script fetches information before it acts on the data.</span></span> <span data-ttu-id="3eb98-114">Например, приведенный ниже код будет точно регистрировать количество строк в используемом диапазоне.</span><span class="sxs-lookup"><span data-stu-id="3eb98-114">For example, the following code will accurately log the number of rows in the used range.</span></span>

```TypeScript
let usedRange = workbook.getActiveWorksheet().getUsedRange();
let rowCount = usedRange.getRowCount();
// The script will read the range and row count from
// the workbook before logging the information.
console.log(rowCount);
```

<span data-ttu-id="3eb98-115">API сценариев Office Обеспечьте точность и актуальность данных в книге или сценарии при необходимости.</span><span class="sxs-lookup"><span data-stu-id="3eb98-115">Office Scripts APIs ensure any data in the workbook or script is accurate and up-to-date when necessary.</span></span> <span data-ttu-id="3eb98-116">Для правильной работы сценария не нужно беспокоиться об этих синхронизациях.</span><span class="sxs-lookup"><span data-stu-id="3eb98-116">You don't need to worry about these synchronizations for your script to run correctly.</span></span> <span data-ttu-id="3eb98-117">Тем не менее, сведения о связи между этими сценариями и облаками могут помочь вам избежать ненужных сетевых вызовов.</span><span class="sxs-lookup"><span data-stu-id="3eb98-117">However, an awareness of this script-to-cloud communication can help you avoid unneeded network calls.</span></span>

## <a name="performance-optimizations"></a><span data-ttu-id="3eb98-118">Оптимизация производительности</span><span class="sxs-lookup"><span data-stu-id="3eb98-118">Performance optimizations</span></span>

<span data-ttu-id="3eb98-119">Вы можете применять простые методы для сокращения взаимодействия с облаком.</span><span class="sxs-lookup"><span data-stu-id="3eb98-119">You can apply simple techniques to help reduce the communication to the cloud.</span></span> <span data-ttu-id="3eb98-120">Следующие шаблоны помогают ускорить работу ваших сценариев.</span><span class="sxs-lookup"><span data-stu-id="3eb98-120">The following patterns help speed up your scripts.</span></span>

- <span data-ttu-id="3eb98-121">Считывание данных книги один раз, а не повторно в цикле.</span><span class="sxs-lookup"><span data-stu-id="3eb98-121">Read workbook data once instead of repeatedly in a loop.</span></span>
- <span data-ttu-id="3eb98-122">Удалите лишние `console.log` операторы.</span><span class="sxs-lookup"><span data-stu-id="3eb98-122">Remove unnecessary `console.log` statements.</span></span>
- <span data-ttu-id="3eb98-123">Не используйте блоки try/catch.</span><span class="sxs-lookup"><span data-stu-id="3eb98-123">Avoid using try/catch blocks.</span></span>

### <a name="read-workbook-data-outside-of-a-loop"></a><span data-ttu-id="3eb98-124">Считывание данных книги вне цикла</span><span class="sxs-lookup"><span data-stu-id="3eb98-124">Read workbook data outside of a loop</span></span>

<span data-ttu-id="3eb98-125">Любой метод, который получает данные из книги, может активировать сетевой вызов.</span><span class="sxs-lookup"><span data-stu-id="3eb98-125">Any method that gets data from the workbook can trigger a network call.</span></span> <span data-ttu-id="3eb98-126">Вместо многократного совершения такого же вызова следует по возможности сохранять данные локально.</span><span class="sxs-lookup"><span data-stu-id="3eb98-126">Rather than repeatedly making the same call, you should save data locally whenever possible.</span></span> <span data-ttu-id="3eb98-127">Это особенно справедливо при работе с циклами.</span><span class="sxs-lookup"><span data-stu-id="3eb98-127">This is especially true when dealing with loops.</span></span>

<span data-ttu-id="3eb98-128">Рассмотрим сценарий, чтобы получить количество отрицательных чисел в используемом диапазоне листа.</span><span class="sxs-lookup"><span data-stu-id="3eb98-128">Consider a script to get the count of negative numbers in the used range of a worksheet.</span></span> <span data-ttu-id="3eb98-129">Скрипту необходимо выполнить итерацию по каждой ячейке в используемом диапазоне.</span><span class="sxs-lookup"><span data-stu-id="3eb98-129">The script needs to iterate over every cell in the used range.</span></span> <span data-ttu-id="3eb98-130">Для этого требуется диапазон, количество строк и число столбцов.</span><span class="sxs-lookup"><span data-stu-id="3eb98-130">To do that, it needs the range, the number of rows, and the number of columns.</span></span> <span data-ttu-id="3eb98-131">Перед началом цикла необходимо сохранить их как локальные переменные.</span><span class="sxs-lookup"><span data-stu-id="3eb98-131">You should store those as local variables before starting the loop.</span></span> <span data-ttu-id="3eb98-132">В противном случае каждая итерация цикла будет принудительно возвращаться к книге.</span><span class="sxs-lookup"><span data-stu-id="3eb98-132">Otherwise, each iteration of the loop will force a return to the workbook.</span></span>

```TypeScript
/**
 * This script provides the count of negative numbers that are present
 * in the used range of the current worksheet.
 */
function main(workbook: ExcelScript.Workbook) {
  // Get the working range.
  let usedRange = workbook.getActiveWorksheet().getUsedRange();

  // Save the values locally to avoid repeatedly asking the workbook.
  let usedRangeValues = usedRange.getValues();

  // Start the negative number counter.
  let negativeCount = 0;

  // Iterate over the entire range looking for negative numbers.
  for (let i = 0; i < usedRangeValues.length; i++) {
    for (let j = 0; j < usedRangeValues[i].length; j++) {
      if (usedRangeValues[i][j] < 0) {
        negativeCount++;
      }
    }
  }

  // Log the negative number count to the console.
  console.log(negativeCount);
}
```

> [!NOTE]
> <span data-ttu-id="3eb98-133">В качестве эксперимента попробуйте заменить `usedRangeValues` в цикле `usedRange.getValues()` .</span><span class="sxs-lookup"><span data-stu-id="3eb98-133">As an experiment, try replacing `usedRangeValues` in the loop with `usedRange.getValues()`.</span></span> <span data-ttu-id="3eb98-134">Вы можете заметить, что сценарий выполняется значительно дольше при работе с большими диапазонами.</span><span class="sxs-lookup"><span data-stu-id="3eb98-134">You may notice the script takes considerably longer to run when dealing with large ranges.</span></span>

### <a name="remove-unnecessary-consolelog-statements"></a><span data-ttu-id="3eb98-135">Удаление ненужных `console.log` операторов</span><span class="sxs-lookup"><span data-stu-id="3eb98-135">Remove unnecessary `console.log` statements</span></span>

<span data-ttu-id="3eb98-136">Ведение журнала консоли — это крайне важное средство для [отладки сценариев](../testing/troubleshooting.md).</span><span class="sxs-lookup"><span data-stu-id="3eb98-136">Console logging is a vital tool for [debugging your scripts](../testing/troubleshooting.md).</span></span> <span data-ttu-id="3eb98-137">Тем не менее, он принудительно выполняет синхронизацию с книгой, чтобы убедиться в том, что зарегистрированные данные актуальны.</span><span class="sxs-lookup"><span data-stu-id="3eb98-137">However, it does force the script to synchronize with the workbook to ensure the logged information is up-to-date.</span></span> <span data-ttu-id="3eb98-138">Перед предоставлением общего доступа к сценарию рекомендуется удалить лишние операторы ведения журнала (например, используемые для тестирования).</span><span class="sxs-lookup"><span data-stu-id="3eb98-138">Consider removing unnecessary logging statements (such as those used for testing) before sharing your script.</span></span> <span data-ttu-id="3eb98-139">Это, как правило, не вызывает значительных проблем с производительностью, если `console.log()` оператор не находится в цикле.</span><span class="sxs-lookup"><span data-stu-id="3eb98-139">This typically won't cause a noticeable performance issue, unless the `console.log()` statement is in a loop.</span></span>

### <a name="avoid-using-trycatch-blocks"></a><span data-ttu-id="3eb98-140">Не используйте блоки try/catch</span><span class="sxs-lookup"><span data-stu-id="3eb98-140">Avoid using try/catch blocks</span></span>

<span data-ttu-id="3eb98-141">Мы не рекомендуем использовать [ `try` / `catch` блоки](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) как часть ожидаемого потокового управления сценария.</span><span class="sxs-lookup"><span data-stu-id="3eb98-141">We don't recommend using [`try`/`catch` blocks](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) as part of a script's expected control flow.</span></span> <span data-ttu-id="3eb98-142">Большинство ошибок можно избежать, проверив объекты, возвращенные из книги.</span><span class="sxs-lookup"><span data-stu-id="3eb98-142">Most errors can be avoided by checking objects returned from the workbook.</span></span> <span data-ttu-id="3eb98-143">Например, следующий сценарий проверяет, существует ли таблица, возвращаемая книгой, прежде чем пытаться добавить строку.</span><span class="sxs-lookup"><span data-stu-id="3eb98-143">For example, the following script checks that the table returned by the workbook exists before trying to add a row.</span></span>

```TypeScript
/**
 * This script adds a row to "MyTable", if that table is present.
 */
function main(workbook: ExcelScript.Workbook) {
  let table = workbook.getTable("MyTable");

  // Check if the table exists.
  if (table) {
    // Add the row.
    table.addRow(-1, ["2012", "Yes", "Maybe"]);
  } else {
    // Report the missing table.
    console.log("MyTable not found.");
  }
}
```

## <a name="case-by-case-help"></a><span data-ttu-id="3eb98-144">Справка по случаю обращения</span><span class="sxs-lookup"><span data-stu-id="3eb98-144">Case-by-case help</span></span>

<span data-ttu-id="3eb98-145">Так как платформа сценариев Office расширяется для работы со средствами [автоматизации](https://flow.microsoft.com/), [адаптивными картами](https://docs.microsoft.com/adaptive-cards)и другими компонентами, подробные сведения о связи между книгой становятся более сложными.</span><span class="sxs-lookup"><span data-stu-id="3eb98-145">As the Office Scripts platform expands to work with [Power Automate](https://flow.microsoft.com/), [Adaptive Cards](https://docs.microsoft.com/adaptive-cards), and other cross-product features, the details of the script-workbook communication become more intricate.</span></span> <span data-ttu-id="3eb98-146">Если вам нужна помощь быстрее при выполнении скрипта, обратитесь к разделу [переполнение стека](https://stackoverflow.com/questions/tagged/office-scripts).</span><span class="sxs-lookup"><span data-stu-id="3eb98-146">If you need help making your script run faster, please reach out through [Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts).</span></span> <span data-ttu-id="3eb98-147">Не забудьте пометить свой вопрос знакомым "Office — Scripts", чтобы специалисты могли найти их и помочь.</span><span class="sxs-lookup"><span data-stu-id="3eb98-147">Be sure to tag your question with "office-scripts" so experts can find it and help.</span></span>

## <a name="see-also"></a><span data-ttu-id="3eb98-148">См. также</span><span class="sxs-lookup"><span data-stu-id="3eb98-148">See also</span></span>

- [<span data-ttu-id="3eb98-149">Основы сценариев для сценариев Office в Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="3eb98-149">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
- [<span data-ttu-id="3eb98-150">Веб-документы МДН: циклы и итерация</span><span class="sxs-lookup"><span data-stu-id="3eb98-150">MDN web docs: Loops and iteration</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)
