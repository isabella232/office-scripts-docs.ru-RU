---
title: Чтение данных книги с помощью сценариев Office в Excel в Интернете
description: Учебник по сценариям Office о чтении данных из книг и их оценке в сценарии.
ms.date: 04/23/2020
localization_priority: Priority
ms.openlocfilehash: 93204184d4b5947b2a67107b1fd73c178a73c32e
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878690"
---
# <a name="read-workbook-data-with-office-scripts-in-excel-on-the-web"></a><span data-ttu-id="92679-103">Чтение данных книги с помощью сценариев Office в Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="92679-103">Read workbook data with Office Scripts in Excel on the web</span></span>

<span data-ttu-id="92679-104">В этом учебнике объясняется, как читать данные из книги с помощью сценария Office для Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="92679-104">This tutorial teaches you how to read data from a workbook with an Office Script for Excel on the web.</span></span> <span data-ttu-id="92679-105">После этого вы сможете отредактировать прочитанные данные и вернуть их в книгу.</span><span class="sxs-lookup"><span data-stu-id="92679-105">You'll then edit the data you read and put it back in the workbook.</span></span>

> [!TIP]
> <span data-ttu-id="92679-106">Если вы только приступили к работе со сценариями Office, рекомендуем начать с учебника [Запись, редактирование и создание сценариев Office в Excel в Интернете](excel-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="92679-106">If you are new to Office Scripts, we recommend starting with the [Record, edit, and create Office Scripts in Excel on the web](excel-tutorial.md) tutorial.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="92679-107">Необходимые действия</span><span class="sxs-lookup"><span data-stu-id="92679-107">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/tutorial-prerequisites.md)]

> [!IMPORTANT]
> <span data-ttu-id="92679-108">Этот учебник предназначен для пользователей с начальным и средним уровнем знаний по JavaScript или TypeScript.</span><span class="sxs-lookup"><span data-stu-id="92679-108">This tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript.</span></span> <span data-ttu-id="92679-109">Если вы впервые работаете с JavaScript, рекомендуем прочесть [учебник Mozilla по JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span><span class="sxs-lookup"><span data-stu-id="92679-109">If you're new to JavaScript, we recommend reviewing the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span> <span data-ttu-id="92679-110">Чтобы получить дополнительные сведения о среде сценариев, ознакомьтесь со статьей [Сценарии Office в Excel в Интернете](../overview/excel.md).</span><span class="sxs-lookup"><span data-stu-id="92679-110">Visit [Office Scripts in Excel on the web](../overview/excel.md) to learn more about the script environment.</span></span>

## <a name="read-a-cell"></a><span data-ttu-id="92679-111">Чтение ячейки</span><span class="sxs-lookup"><span data-stu-id="92679-111">Read a cell</span></span>

<span data-ttu-id="92679-112">Сценарии, созданные с помощью средства записи действий, могут только записывать информацию в книгу.</span><span class="sxs-lookup"><span data-stu-id="92679-112">Scripts made with the Action Recorder can only write information to the workbook.</span></span> <span data-ttu-id="92679-113">С помощью редактора кода можно редактировать и создавать сценарии, которые также читают данные из книги.</span><span class="sxs-lookup"><span data-stu-id="92679-113">With the Code Editor, you can edit and make scripts that also read data from a workbook.</span></span>

<span data-ttu-id="92679-114">Давайте создадим сценарий, читающий данные и действующий на основе прочитанного.</span><span class="sxs-lookup"><span data-stu-id="92679-114">Let's make a script that reads data and acts based on what was read.</span></span> <span data-ttu-id="92679-115">Мы будем работать с примером банковской выписки.</span><span class="sxs-lookup"><span data-stu-id="92679-115">We're going to work with a sample banking statement.</span></span> <span data-ttu-id="92679-116">Эта выписка объединяет чековую выписку и выписку по кредиту.</span><span class="sxs-lookup"><span data-stu-id="92679-116">This statement is a combined checking and credit statement.</span></span> <span data-ttu-id="92679-117">К сожалению, изменения баланса в них указываются по-разному.</span><span class="sxs-lookup"><span data-stu-id="92679-117">Unfortunately, they report balance changes differently.</span></span> <span data-ttu-id="92679-118">В чековой выписке доходы указываются как положительный кредит, а расходы — в виде отрицательного дебета.</span><span class="sxs-lookup"><span data-stu-id="92679-118">The checking statement gives income as positive credit and costs as negative debit.</span></span> <span data-ttu-id="92679-119">В выписке по кредиту все наоборот.</span><span class="sxs-lookup"><span data-stu-id="92679-119">The credit statement does the opposite.</span></span>

<span data-ttu-id="92679-120">В остальной части учебника мы нормализуем эти данные с помощью сценария.</span><span class="sxs-lookup"><span data-stu-id="92679-120">Over the rest of the tutorial, we will normalize this data using a script.</span></span> <span data-ttu-id="92679-121">Сначала научимся читать данные из книги.</span><span class="sxs-lookup"><span data-stu-id="92679-121">First, let's learn how to read data from the workbook.</span></span>

1. <span data-ttu-id="92679-122">Создайте лист в книге, которую вы использовали в остальной части учебника.</span><span class="sxs-lookup"><span data-stu-id="92679-122">Create a new worksheet in the workbook you've used for the rest of the tutorial.</span></span>
2. <span data-ttu-id="92679-123">Скопируйте следующие данные и вставьте их на новый лист, начиная с ячейки **A1**.</span><span class="sxs-lookup"><span data-stu-id="92679-123">Copy the following data and paste it into the new worksheet, starting at cell **A1**.</span></span>

    |<span data-ttu-id="92679-124">Дата</span><span class="sxs-lookup"><span data-stu-id="92679-124">Date</span></span> |<span data-ttu-id="92679-125">Счет</span><span class="sxs-lookup"><span data-stu-id="92679-125">Account</span></span> |<span data-ttu-id="92679-126">Описание</span><span class="sxs-lookup"><span data-stu-id="92679-126">Description</span></span> |<span data-ttu-id="92679-127">Дебет</span><span class="sxs-lookup"><span data-stu-id="92679-127">Debit</span></span> |<span data-ttu-id="92679-128">Кредит</span><span class="sxs-lookup"><span data-stu-id="92679-128">Credit</span></span> |
    |:--|:--|:--|:--|:--|
    |<span data-ttu-id="92679-129">10.10.2019</span><span class="sxs-lookup"><span data-stu-id="92679-129">10/10/2019</span></span> |<span data-ttu-id="92679-130">Чековый</span><span class="sxs-lookup"><span data-stu-id="92679-130">Checking</span></span> |<span data-ttu-id="92679-131">Виноградник Coho</span><span class="sxs-lookup"><span data-stu-id="92679-131">Coho Vineyard</span></span> |<span data-ttu-id="92679-132">–20,05</span><span class="sxs-lookup"><span data-stu-id="92679-132">-20.05</span></span> | |
    |<span data-ttu-id="92679-133">11.10.2019</span><span class="sxs-lookup"><span data-stu-id="92679-133">10/11/2019</span></span> |<span data-ttu-id="92679-134">Кредитный</span><span class="sxs-lookup"><span data-stu-id="92679-134">Credit</span></span> |<span data-ttu-id="92679-135">Телефонная компания</span><span class="sxs-lookup"><span data-stu-id="92679-135">The Phone Company</span></span> |<span data-ttu-id="92679-136">99,95</span><span class="sxs-lookup"><span data-stu-id="92679-136">99.95</span></span> | |
    |<span data-ttu-id="92679-137">13.10.2019</span><span class="sxs-lookup"><span data-stu-id="92679-137">10/13/2019</span></span> |<span data-ttu-id="92679-138">Кредитный</span><span class="sxs-lookup"><span data-stu-id="92679-138">Credit</span></span> |<span data-ttu-id="92679-139">Виноградник Coho</span><span class="sxs-lookup"><span data-stu-id="92679-139">Coho Vineyard</span></span> |<span data-ttu-id="92679-140">154,43</span><span class="sxs-lookup"><span data-stu-id="92679-140">154.43</span></span> | |
    |<span data-ttu-id="92679-141">15.10.2019</span><span class="sxs-lookup"><span data-stu-id="92679-141">10/15/2019</span></span> |<span data-ttu-id="92679-142">Чековый</span><span class="sxs-lookup"><span data-stu-id="92679-142">Checking</span></span> |<span data-ttu-id="92679-143">Внешний депозит</span><span class="sxs-lookup"><span data-stu-id="92679-143">External Deposit</span></span> | |<span data-ttu-id="92679-144">1000</span><span class="sxs-lookup"><span data-stu-id="92679-144">1000</span></span> |
    |<span data-ttu-id="92679-145">20.10.2019</span><span class="sxs-lookup"><span data-stu-id="92679-145">10/20/2019</span></span> |<span data-ttu-id="92679-146">Кредитный</span><span class="sxs-lookup"><span data-stu-id="92679-146">Credit</span></span> |<span data-ttu-id="92679-147">Виноградник Coho — возмещение</span><span class="sxs-lookup"><span data-stu-id="92679-147">Coho Vineyard - Refund</span></span> | |<span data-ttu-id="92679-148">–35,45</span><span class="sxs-lookup"><span data-stu-id="92679-148">-35.45</span></span> |
    |<span data-ttu-id="92679-149">25.10.2019</span><span class="sxs-lookup"><span data-stu-id="92679-149">10/25/2019</span></span> |<span data-ttu-id="92679-150">Чековый</span><span class="sxs-lookup"><span data-stu-id="92679-150">Checking</span></span> |<span data-ttu-id="92679-151">Органическая компания "Лучшее для вас"</span><span class="sxs-lookup"><span data-stu-id="92679-151">Best For You Organics Company</span></span> | <span data-ttu-id="92679-152">–85,64</span><span class="sxs-lookup"><span data-stu-id="92679-152">-85.64</span></span> | |
    |<span data-ttu-id="92679-153">01.11.2019</span><span class="sxs-lookup"><span data-stu-id="92679-153">11/01/2019</span></span> |<span data-ttu-id="92679-154">Чековый</span><span class="sxs-lookup"><span data-stu-id="92679-154">Checking</span></span> |<span data-ttu-id="92679-155">Внешний депозит</span><span class="sxs-lookup"><span data-stu-id="92679-155">External Deposit</span></span> | |<span data-ttu-id="92679-156">1000</span><span class="sxs-lookup"><span data-stu-id="92679-156">1000</span></span> |

3. <span data-ttu-id="92679-157">Откройте **Редактор кода** и выберите **Создать сценарий**.</span><span class="sxs-lookup"><span data-stu-id="92679-157">Open the **Code Editor** and select **New Script**.</span></span>
4. <span data-ttu-id="92679-158">Давайте очистим форматирование.</span><span class="sxs-lookup"><span data-stu-id="92679-158">Let's clean up the formatting.</span></span> <span data-ttu-id="92679-159">Это финансовый документ, поэтому изменим числовой формат в столбцах **Дебет** и **Кредит**, чтобы отобразить значения в долларах.</span><span class="sxs-lookup"><span data-stu-id="92679-159">This is a financial document, so let's change the number formatting in the **Debit** and **Credit** columns to show values as dollar amounts.</span></span> <span data-ttu-id="92679-160">Также настроим ширину столбца по данным.</span><span class="sxs-lookup"><span data-stu-id="92679-160">Let's also fit the column width to the data.</span></span>

    <span data-ttu-id="92679-161">Замените содержимое сценария следующим кодом:</span><span class="sxs-lookup"><span data-stu-id="92679-161">Replace the script contents with the following code:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
        // Get the current worksheet.
        let selectedSheet = workbook.getActiveWorksheet();

        // Format the range to display numerical dollar amounts.
        selectedSheet.getRange("D2:E8").setNumberFormat("$#,##0.00");

        // Fit the width of all the used columns to the data.
        selectedSheet.getUsedRange().getFormat().autofitColumns();
    }
    ```

5. <span data-ttu-id="92679-162">Теперь прочитаем значение в одном из числовых столбцов.</span><span class="sxs-lookup"><span data-stu-id="92679-162">Now let's read a value from one of the number columns.</span></span> <span data-ttu-id="92679-163">Добавьте следующий код в конце сценария (перед закрывающей скобкой `}`):</span><span class="sxs-lookup"><span data-stu-id="92679-163">Add the following code to the end of the script (before the closing `}`):</span></span>

    ```TypeScript
    // Get the value of cell D2.
    let range = selectedSheet.getRange("D2");
    console.log(range.getValues());
    ```

6. <span data-ttu-id="92679-164">Запустите сценарий.</span><span class="sxs-lookup"><span data-stu-id="92679-164">Run the script.</span></span>
7. <span data-ttu-id="92679-165">Откройте консоль.</span><span class="sxs-lookup"><span data-stu-id="92679-165">Open the console.</span></span> <span data-ttu-id="92679-166">Откройте меню **Многоточие** и нажмите **Журналы...**.</span><span class="sxs-lookup"><span data-stu-id="92679-166">Go to the **Ellipses** menu and press **Logs...**.</span></span>
8. <span data-ttu-id="92679-167">В консоли должно отображаться следующее: `[Array[1]]`.</span><span class="sxs-lookup"><span data-stu-id="92679-167">You should see `[Array[1]]` in the console.</span></span> <span data-ttu-id="92679-168">Это не число, так как диапазоны являются двухмерными массивами данных.</span><span class="sxs-lookup"><span data-stu-id="92679-168">This is not a number because ranges are two-dimensional arrays of data.</span></span> <span data-ttu-id="92679-169">Этот двухмерный диапазон напрямую регистрируется в консоли.</span><span class="sxs-lookup"><span data-stu-id="92679-169">That two-dimensional range is being logged to the console directly.</span></span> <span data-ttu-id="92679-170">К счастью, редактор кода позволяет просмотреть содержимое массива.</span><span class="sxs-lookup"><span data-stu-id="92679-170">Luckily, the Code Editor does let you see the contents of the array.</span></span>
9. <span data-ttu-id="92679-171">Когда двухмерный массив регистрируется в консоли, она группирует значения столбцов под каждой строкой.</span><span class="sxs-lookup"><span data-stu-id="92679-171">When a two-dimensional array is logged to the console, it groups column values under each row.</span></span> <span data-ttu-id="92679-172">Разверните журнал массива, нажав синий треугольник.</span><span class="sxs-lookup"><span data-stu-id="92679-172">Expand the array log by pressing the blue triangle.</span></span>
10. <span data-ttu-id="92679-173">Разверните второй уровень массива, нажав появившийся синий треугольник.</span><span class="sxs-lookup"><span data-stu-id="92679-173">Expand the second level of the array by pressing the newly revealed blue triangle.</span></span> <span data-ttu-id="92679-174">Должно отобразиться следующее:</span><span class="sxs-lookup"><span data-stu-id="92679-174">You should now see this:</span></span>

    ![Журнал консоли, отображающий результат "–20,05", размещенный под двумя массивами.](../images/tutorial-4.png)

## <a name="modify-the-value-of-a-cell"></a><span data-ttu-id="92679-176">Изменение значения ячейки</span><span class="sxs-lookup"><span data-stu-id="92679-176">Modify the value of a cell</span></span>

<span data-ttu-id="92679-177">Теперь, когда мы можем читать данные, воспользуемся ими, чтобы изменить книгу.</span><span class="sxs-lookup"><span data-stu-id="92679-177">Now that we can read data, let's use that data to modify the workbook.</span></span> <span data-ttu-id="92679-178">Мы сделаем значение ячейки **D2** положительным с помощью функции `Math.abs`.</span><span class="sxs-lookup"><span data-stu-id="92679-178">We'll make the value of the cell **D2** positive with the `Math.abs` function.</span></span> <span data-ttu-id="92679-179">Объект [Math](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/math) содержит множество функций, к которым имеют доступ сценарии.</span><span class="sxs-lookup"><span data-stu-id="92679-179">The [Math](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/math) object contains many functions to which your scripts have access.</span></span> <span data-ttu-id="92679-180">Дополнительные сведения о `Math` и других встроенных объектах можно найти в статье [Использование встроенных объектов JavaScript в сценариях Office](../develop/javascript-objects.md).</span><span class="sxs-lookup"><span data-stu-id="92679-180">More information about `Math` and other built-in objects can be found at [Using built-in JavaScript objects in Office Scripts](../develop/javascript-objects.md).</span></span>

1. <span data-ttu-id="92679-181">Добавьте следующий код в конце сценария:</span><span class="sxs-lookup"><span data-stu-id="92679-181">Add the following code to the end of the script:</span></span>

    ```TypeScript
        // Run the `Math.abs` function with the value at D2 and apply that value back to D2.
    let positiveValue = Math.abs(range.getValue());
    range.setValue(positiveValue);
    ```

    <span data-ttu-id="92679-182">Обратите внимание на то, что мы используем `getValue` и `setValue`.</span><span class="sxs-lookup"><span data-stu-id="92679-182">Note that we're using `getValue` and `setValue`.</span></span> <span data-ttu-id="92679-183">Эти методы применимы к одной ячейке.</span><span class="sxs-lookup"><span data-stu-id="92679-183">These methods work on a single cell.</span></span> <span data-ttu-id="92679-184">При обработке диапазонов, включающих несколько ячеек, нужно использовать `getValues` и `setValues`.</span><span class="sxs-lookup"><span data-stu-id="92679-184">When handling multi-cell ranges, you'll want to use `getValues` and `setValues`.</span></span>

2. <span data-ttu-id="92679-185">Значение ячейки **D2** теперь должно быть положительным.</span><span class="sxs-lookup"><span data-stu-id="92679-185">The value of cell **D2** should now be positive.</span></span>

## <a name="modify-the-values-of-a-column"></a><span data-ttu-id="92679-186">Изменение значений столбца</span><span class="sxs-lookup"><span data-stu-id="92679-186">Modify the values of a column</span></span>

<span data-ttu-id="92679-187">Теперь, когда вы знаете, как читать и записывать данные в одной ячейке, давайте обобщим сценарий для работы со всеми столбцами **Дебет** и **Кредит**.</span><span class="sxs-lookup"><span data-stu-id="92679-187">Now that we know how to read and write to a single cell, let's generalize the script to work on the entire **Debit** and **Credit** columns.</span></span>

1. <span data-ttu-id="92679-188">Удалите код, влияющий только на одну ячейку (предыдущий код с абсолютным значением), чтобы ваш сценарий выглядел следующим образом:</span><span class="sxs-lookup"><span data-stu-id="92679-188">Remove the code that affects only a single cell (the previous absolute value code), such that your script now looks like this:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
        // Get the current worksheet.
        let selectedSheet = workbook.getActiveWorksheet();

        // Format the range to display numerical dollar amounts.
        selectedSheet.getRange("D2:E8").setNumberFormat("$#,##0.00");

        // Fit the width of all the used columns to the data.
        selectedSheet.getUsedRange().getFormat().autofitColumns();
    }
    ```

2. <span data-ttu-id="92679-189">Добавьте цикл в конце сценария, выполняющий итерацию в строках двух последних столбцов.</span><span class="sxs-lookup"><span data-stu-id="92679-189">Add a loop to the end of the script that iterates through the rows in the last two columns.</span></span> <span data-ttu-id="92679-190">Для каждой ячейки сценарий устанавливает текущее абсолютное значение.</span><span class="sxs-lookup"><span data-stu-id="92679-190">For each cell, the script sets the value to the current value's absolute value.</span></span>

    <span data-ttu-id="92679-191">Обратите внимание, что массив, определяющий расположения ячеек, отсчитывается от нуля.</span><span class="sxs-lookup"><span data-stu-id="92679-191">Note that the array defining cell locations is zero-based.</span></span> <span data-ttu-id="92679-192">Это означает, что ячейка **A1** имеет значение `range[0][0]`.</span><span class="sxs-lookup"><span data-stu-id="92679-192">That means cell **A1** is `range[0][0]`.</span></span>

    ```TypeScript
    // Get the values of the used range.
    let range = selectedSheet.getUsedRange();
    let rangeValues = range.getValues();

    // Iterate over the fourth and fifth columns and set their values to their absolute value.
    for (let i = 1; i < range.getRowCount(); i++) {
        // The column at index 3 is column "4" in the worksheet.
        if (rangeValues[i][3] != 0) {
            let positiveValue = Math.abs(rangeValues[i][3]);
            selectedSheet.getCell(i, 3).setValue(positiveValue);
        }

        // The column at index 4 is column "5" in the worksheet.
        if (rangeValues[i][4] != 0) {
            let positiveValue = Math.abs(rangeValues[i][4]);
            selectedSheet.getCell(i, 4).setValue(positiveValue);
        }
    }
    ```

    <span data-ttu-id="92679-193">Эта часть сценария выполняет несколько важных задач.</span><span class="sxs-lookup"><span data-stu-id="92679-193">This portion of the script does several important tasks.</span></span> <span data-ttu-id="92679-194">Сначала она получает значения и количество строк используемого диапазона.</span><span class="sxs-lookup"><span data-stu-id="92679-194">First, it gets the values and row count of the used range.</span></span> <span data-ttu-id="92679-195">Это позволяет просматривать значения и определять момент остановки.</span><span class="sxs-lookup"><span data-stu-id="92679-195">This lets us look at values and know when to stop.</span></span> <span data-ttu-id="92679-196">Затем выполняется итерация в используемом диапазоне с проверкой каждой ячейки в столбцах **Дебет** или **Кредит**.</span><span class="sxs-lookup"><span data-stu-id="92679-196">Second, it iterates through the used range, checking each cell in the **Debit** or **Credit** columns.</span></span> <span data-ttu-id="92679-197">Наконец, если значение в ячейке не равно 0, оно заменяется абсолютным значением.</span><span class="sxs-lookup"><span data-stu-id="92679-197">Finally, if the value in the cell is not 0, it is replaced by its absolute value.</span></span> <span data-ttu-id="92679-198">Мы избегаем использования нулей, поэтому можно оставить пустые ячейки неизменными.</span><span class="sxs-lookup"><span data-stu-id="92679-198">We're avoiding zeroes so we can leave the blank cells as they were.</span></span>

3. <span data-ttu-id="92679-199">Запустите сценарий.</span><span class="sxs-lookup"><span data-stu-id="92679-199">Run the script.</span></span>

    <span data-ttu-id="92679-200">Теперь банковская выписка должна выглядеть следующим образом:</span><span class="sxs-lookup"><span data-stu-id="92679-200">Your banking statement should now look like this:</span></span>

    ![Банковская выписка в виде отформатированной таблицы только с положительными значениями.](../images/tutorial-5.png)

## <a name="next-steps"></a><span data-ttu-id="92679-202">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="92679-202">Next steps</span></span>

<span data-ttu-id="92679-203">Откройте редактор кода и попробуйте некоторые [примеры сценариев Office в Excel в Интернете](../resources/excel-samples.md).</span><span class="sxs-lookup"><span data-stu-id="92679-203">Open the Code Editor and try out some of our [Sample scripts for Office Scripts in Excel on the web](../resources/excel-samples.md).</span></span> <span data-ttu-id="92679-204">Дополнительные сведения о создании сценариев Office доступны в статье [Основные сведения о сценариях Office в Excel в Интернете](../develop/scripting-fundamentals.md).</span><span class="sxs-lookup"><span data-stu-id="92679-204">You can also visit [Scripting Fundamentals for Office Scripts in Excel on the web](../develop/scripting-fundamentals.md) to learn more about creating Office Scripts.</span></span>
